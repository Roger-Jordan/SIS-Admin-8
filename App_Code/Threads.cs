using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.Text.RegularExpressions;
using System.Configuration;

/// <summary>
/// Zusammenfassungsbeschreibung für PDFThread
/// </summary>
public class Threads
{
    public class TEmployeeData
    {
        public int employeeID;
        public Dictionary<string, string> values;
    }
    protected TSessionObj SessionObj;

    public Threads(TSessionObj aSessionObj)
    {
        SessionObj = aSessionObj;
    }

    //-------------------- Threads für Ingress-Export -------------------------------
    # region IngressExport
    public void doIngressExcelExport()
    {
        // Excel Datei mit 2 Register erzeugen
        ExcelFile tempxls = new XlsFile(true);
        tempxls.NewFile(2, TExcelFileFormat.v2010);

        tempxls.ActiveSheet = 1;
        tempxls.SheetName = "Emloyees Online";
        tempxls.ActiveSheet = 2;
        tempxls.SheetName = "Emloyees Paper";

        TIngressMapping mapping = new TIngressMapping(SessionObj.Project.ProjectID);

        // Header schreiben
        int counterOnline = 1;
        int counterPaper = 1;
        foreach (TIngressMapping.TMappingElement tempMappingElement in mapping.mapping)
        {
            if (tempMappingElement.method == 1)
            {
                tempxls.ActiveSheet = 2;
                counterPaper++;
            }
            else
            {
                tempxls.ActiveSheet = 1;
                counterOnline++;
            }
            tempxls.SetCellValue(1, tempMappingElement.position, tempMappingElement.targetname);
        }
        tempxls.ActiveSheet = 1;
        tempxls.SetCellValue(1, counterOnline, "PostNomination");
        tempxls.ActiveSheet = 2;
        tempxls.SetCellValue(1, counterPaper, "PostNomination");

        // Suchkriterium festlegen
        #region Filter
        string searchString = "";

        if (SessionObj.postnominationExportType == "")
        {
            searchString = "SELECT employeeID, name, firstname, email, orgID, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, issupervisor, issupervisor1, issupervisor2, issupervisor3, mobile, surveytype, accesscode, postnomination  FROM orgmanager_employee";
        }
        else
        {
            #region SurveyType-Filter
            if (SessionObj.postnominationExportGroup.IndexOf("Online") >= 0)
            {
                if (searchString == "")
                    searchString = "surveytype='0'";
                else
                    searchString += " OR surveytype='0'";
            }
            if (SessionObj.postnominationExportGroup.IndexOf("Paper") >= 0)
            {
                if (searchString == "")
                    searchString = "surveytype='1'";
                else
                    searchString += " OR surveytype='1'";
            }
            if (SessionObj.postnominationExportGroup.IndexOf("Hybrid") >= 0)
            {
                if (searchString == "")
                    searchString = "surveytype='2'";
                else
                    searchString += " OR surveytype='2'";
            }
            if (SessionObj.postnominationExportGroup.IndexOf("SMS") >= 0)
            {
                if (searchString == "")
                    searchString = "surveytype='3'";
                else
                    searchString += " OR surveytype='3'";
            }
            #endregion SurveyType-Filter

            if (SessionObj.postnominationExportType == "all-postnominated")
            {
                // Export von Nachnominierungen
                if (searchString == "")
                    searchString = " WHERE postnomination = '" + SessionObj.postnominationExportType + "'";
                else
                    searchString = " WHERE (" + searchString + ") AND  (postnomination != '')";

                searchString = "SELECT employeeID, name, firstname, email, orgID, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, issupervisor, issupervisor1, issupervisor2, issupervisor3, mobile, surveytype, accesscode, postnomination FROM orgmanager_employee" + searchString;
            }

            if (SessionObj.postnominationExportType == "postnominated")
            {
                // Export von Nachnominierungen
                if (searchString == "")
                    searchString = " WHERE postnomination = '" + SessionObj.postnominationExportType + "'";
                else
                    searchString = " WHERE (" + searchString + ") AND  (postnomination = '" + SessionObj.postnominationExportType + "')";

                searchString = "SELECT employeeID, name, firstname, email, orgID, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, issupervisor, issupervisor1, issupervisor2, issupervisor3, mobile, surveytype, accesscode, postnomination FROM orgmanager_employee" + searchString;
            }

            if (SessionObj.postnominationExportType == "exported")
            {
                // Export von Nachnominierungen
                if (searchString == "")
                    searchString = " WHERE postnomination = '" + SessionObj.postnominationExportType + "'";
                else
                    searchString = " WHERE (" + searchString + ") AND  (postnomination = '" + SessionObj.postnominationExportType + "')";

                searchString = "SELECT employeeID, name, firstname, email, orgID, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, issupervisor, issupervisor1, issupervisor2, issupervisor3, mobile, surveytype, accesscode, postnomination FROM orgmanager_employee" + searchString;
            }
        }

        searchString += " ORDER BY employeeID";
        #endregion Filter

        // Schleife über alle relevanten Mitarbeiter
        int actRow = 0;
        int actRowOnline = 2;
        int actRowPaper = 2;
        SqlDB dataReader = new SqlDB(searchString, SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDIngressExport, 4, actRowOnline + actRowPaper - 4, 0, 0, SessionObj.Project.ProjectID);

        int employeeMethod = dataReader.getInt32(17);
            if (employeeMethod == 1)
            {
                tempxls.ActiveSheet = 2;
                actRow = actRowPaper;
                // Postnomination-Status ausgeben
                tempxls.SetCellValue(actRow, counterPaper, dataReader.getString(19));
            }
            else
            {
                tempxls.ActiveSheet = 1;
                actRow = actRowOnline;
                // Postnomination-Status ausgeben
                tempxls.SetCellValue(actRow, counterOnline, dataReader.getString(19));
            }

            // neuer Mitarbeiter
            TEmployeeData actEmployee = new TEmployeeData();
            actEmployee.values = new Dictionary<string, string>();

            // Levelvariablen ermitteln
            int actOrgID = dataReader.getInt32(4);
            ArrayList levels = new ArrayList();
            TAdminStructure.getTopOrgIDList(actOrgID, ref levels, SessionObj.Project.ProjectID);
            int actOrg1ID = dataReader.getInt32(5);
            ArrayList levels1 = new ArrayList();
            TAdminStructure1.getTopOrgIDList(actOrg1ID, ref levels1, SessionObj.Project.ProjectID);
            int actOrg2ID = dataReader.getInt32(6);
            ArrayList levels2 = new ArrayList();
            TAdminStructure2.getTopOrgIDList(actOrg2ID, ref levels2, SessionObj.Project.ProjectID);
            int actOrg3ID = dataReader.getInt32(7);
            ArrayList levels3 = new ArrayList();
            TAdminStructure3.getTopOrgIDList(actOrg3ID, ref levels3, SessionObj.Project.ProjectID);

            actEmployee.employeeID = dataReader.getInt32(0);
            // Schleife über alle Variablen
            foreach (TIngressMapping.TMappingElement actElement in mapping.mapping)
            {
                // Test ob Methode richtige Methode
                if (((actElement.method == 0) && ((employeeMethod == 0) || (employeeMethod == 2) || (employeeMethod == 3))) || ((actElement.method == 1) && (employeeMethod == 1)))
                {
                    switch (actElement.source)
                    {
                        case "Employee":
                            #region Employee
                            // Daten kommen aus der Mitarbeiterliste
                            if (actElement.sourcecolumn == "EmployeeID")
                            {
                                tempxls.SetCellValue(actRow, actElement.position, actEmployee.employeeID);
                                actEmployee.values.Add(actElement.targetname, actEmployee.employeeID.ToString());
                            }
                            else
                            {
                                if (actElement.sourcecolumn == "AccessCode")
                                {
                                    string actAccesscode = dataReader.getString(18);
                                    if (actAccesscode == "")
                                    {
                                        actAccesscode = SessionObj.AccessCode.createAndSaveAccessCode(dataReader.getInt32(0));
                                    }
                                    actEmployee.values.Add("survey_mail_password", actAccesscode);
                                    tempxls.SetCellValue(actRow, actElement.position, actAccesscode);
                                }
                                else
                                {
                                    if (actElement.sourcecolumn == "OrgID")
                                    {
                                        actEmployee.values.Add(actElement.targetname, actOrgID.ToString());
                                        tempxls.SetCellValue(actRow, actElement.position, actOrgID.ToString());
                                    }
                                    else
                                    {
                                        // Datenfeldtyp ermitteln
                                        string actFieldtype = "";
                                        SqlDB dataReader2;
                                        dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_employeefields WHERE fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                        if (dataReader2.read())
                                            actFieldtype = dataReader2.getString(0);
                                        dataReader2.close();

                                        // Wert ermitteln
                                        switch (actFieldtype)
                                        {
                                            case "NameEmployee":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getString(1).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getString(1).ToString());
                                                break;
                                            case "FirstnameEmployee":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getString(2).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getString(2).ToString());
                                                break;
                                            case "EmailEmployee":
                                                if (actElement.method == 3)
                                                {
                                                    actEmployee.values.Add(actElement.targetname, dataReader.getString(16).ToString() + "@sms.messagebird.com");
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader.getString(16).ToString() + "@sms.messagebird.com");
                                                }
                                                else
                                                {
                                                    actEmployee.values.Add(actElement.targetname, dataReader.getString(3).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader.getString(3).ToString());
                                                }
                                                break;
                                            case "MobileEmployee":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getString(16).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getString(16).ToString());
                                                break;
                                            case "UnitManager":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(12).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(12).ToString());
                                                break;
                                            case "UnitManager1":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(13).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(13).ToString());
                                                break;
                                            case "UnitManager2":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(14).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(14).ToString());
                                                break;
                                            case "UnitManager3":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(15).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(15).ToString());
                                                break;
                                            case "SurveyType":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(17).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(17).ToString());
                                                break;
                                            case "Matrix1":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(5).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(5).ToString());
                                                break;
                                            case "Matrix2":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(6).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(6).ToString());
                                                break;
                                            case "Matrix3":
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(7).ToString());
                                                tempxls.SetCellValue(actRow, actElement.position, dataReader.getInt32(7).ToString());
                                                break;
                                            default:
                                                dataReader2 = new SqlDB("Select value FROM orgmanager_employeefieldsvalues WHERE employeeID='" + actEmployee.employeeID + "' AND fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                                if (dataReader2.read())
                                                {
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0));
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(0));
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
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[0]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels.Count > 1)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[1]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[0]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels.Count > 2)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[2]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels.Count > 3)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[3]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels.Count > 4)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[4]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels.Count > 5)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[5]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels.Count > 6)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[6]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels.Count > 7)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[7]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels.Count > 8)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[8]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels.Count > 9)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[9]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels.Count > 10)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[10]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels.Count > 11)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[11]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels.Count > 12)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[12]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels.Count > 13)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[13]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels.Count > 14)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[14]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels.Count > 15)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[15]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels.Count > 16)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[16]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels.Count > 17)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[17]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels.Count > 18)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[18]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels.Count > 19)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[19]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels.Count > 20)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels[20]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structurefields WHERE fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure WHERE orgID='" + actOrgID + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure WHERE orgID='" + actOrgID + "'", SessionObj.Project.ProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(0).ToString());
                                                    break;
                                                case "UnitNameShort":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(1).ToString());
                                                    break;
                                                case "EmployeeCount":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4;
                                                    dataReader4 = new SqlDB("Select value FROM orgmanager_structurefieldsvalues WHERE orgID='" + actOrgID + "' AND fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                        tempxls.SetCellValue(actRow, actElement.position, dataReader4.getString(0));
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
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[0]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels1.Count > 1)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[1]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[0]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels1.Count > 2)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[2]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels1.Count > 3)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[3]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels1.Count > 4)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[4]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels1.Count > 5)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[5]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels1.Count > 6)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[6]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels1.Count > 7)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[7]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels1.Count > 8)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[8]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels1.Count > 9)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[9]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels1.Count > 10)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[10]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels1.Count > 11)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[11]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels1.Count > 12)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[12]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels1.Count > 13)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[13]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels1.Count > 14)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[14]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels1.Count > 15)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[15]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels1.Count > 16)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[16]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels1.Count > 17)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[17]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels1.Count > 18)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[18]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels1.Count > 19)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[19]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels1.Count > 20)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels1[20]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structure1fields WHERE fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure1 WHERE orgID='" + actOrg1ID + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure1 WHERE orgID='" + actOrg1ID + "'", SessionObj.Project.ProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(0).ToString());
                                                    break;
                                                case "UnitNameShort":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(1).ToString());
                                                    break;
                                                case "EmployeeCount":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4;
                                                    dataReader4 = new SqlDB("Select value FROM orgmanager_structure1fieldsvalues WHERE orgID='" + actOrg1ID + "' AND fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                        tempxls.SetCellValue(actRow, actElement.position, dataReader4.getString(0));
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
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[0]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels2.Count > 1)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[1]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[0]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels2.Count > 2)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[2]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels2.Count > 3)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[3]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels2.Count > 4)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[4]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels2.Count > 5)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[5]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels2.Count > 6)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[6]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels2.Count > 7)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[7]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels2.Count > 8)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[8]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels2.Count > 9)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[9]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels2.Count > 10)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[10]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels2.Count > 11)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[11]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels2.Count > 12)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[12]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels2.Count > 13)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[13]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels2.Count > 14)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[14]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels2.Count > 15)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[15]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels2.Count > 16)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[16]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels2.Count > 17)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[17]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels2.Count > 18)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[18]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels2.Count > 19)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[19]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels2.Count > 20)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels2[20]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structure2fields WHERE fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure2 WHERE orgID='" + actOrg2ID + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure2 WHERE orgID='" + actOrg2ID + "'", SessionObj.Project.ProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(0).ToString());
                                                    break;
                                                case "UnitNameShort":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(1).ToString());
                                                    break;
                                                case "EmployeeCount":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4;
                                                    dataReader4 = new SqlDB("Select value FROM orgmanager_structure2fieldsvalues WHERE orgID='" + actOrg2ID + "' AND fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                        tempxls.SetCellValue(actRow, actElement.position, dataReader4.getString(0));
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
                            // Daten kommen aus der Struktur2
                            switch (actElement.sourcecolumn)
                            {
                                case "Level0":
                                    if (levels3.Count > 0)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[0]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels3.Count > 1)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[1]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[0]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels3.Count > 2)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[2]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels3.Count > 3)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[3]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels3.Count > 4)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[4]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels3.Count > 5)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[5]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels3.Count > 6)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[6]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels3.Count > 7)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[7]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels3.Count > 8)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[8]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels3.Count > 9)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[9]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels3.Count > 10)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[10]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels3.Count > 11)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[11]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels3.Count > 12)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[12]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels3.Count > 13)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[13]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels3.Count > 14)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[14]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels3.Count > 15)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[15]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels3.Count > 16)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[16]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels3.Count > 17)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[17]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels3.Count > 18)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[18]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels3.Count > 19)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[19]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels3.Count > 20)
                                    {
                                        tempxls.SetCellValue(actRow, actElement.position, (int)levels3[20]);
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structure3fields WHERE fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure3 WHERE orgID='" + actOrg3ID + "'", SessionObj.Project.ProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure3 WHERE orgID='" + actOrg3ID + "'", SessionObj.Project.ProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(0).ToString());
                                                    break;
                                                case "UnitNameShort":
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader2.getString(1).ToString());
                                                    break;
                                                case "EmployeeCount":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    tempxls.SetCellValue(actRow, actElement.position, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4;
                                                    dataReader4 = new SqlDB("Select value FROM orgmanager_structure3fieldsvalues WHERE orgID='" + actOrg3ID + "' AND fieldID='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                        tempxls.SetCellValue(actRow, actElement.position, dataReader4.getString(0));
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
                            string tempValue = "### Missing Value ###";
                            SqlDB dataReader1;
                            dataReader1 = new SqlDB("Select value FROM orgmanager_ingressvalues WHERE source='" + actElement.source + "' AND key1='" + key1value + "' AND key2='" + key2value + "' AND key3='" + key3value + "' AND key4='" + key4value + "' AND key5='" + key5value + "' AND valuecolumn='" + actElement.sourcecolumn + "'", SessionObj.Project.ProjectID);
                            if (dataReader1.read())
                            {
                                tempValue = dataReader1.getString(0);
                            }
                            else
                            {
                                //### Log-Eintrag in separates Register



                            }
                            dataReader1.close();
                            actEmployee.values.Add(actElement.targetname, tempValue);
                            tempxls.SetCellValue(actRow, actElement.position, tempValue);
                            break;
                            #endregion Zusatzwerte
                    }
                }
            }

            if (employeeMethod == 1)
                actRowPaper++;
            else
                actRowOnline++;
        }
        dataReader.close();

        // Excel generieren
        SessionObj.ProcessMessage = "Excel streamen ... ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDIngressExport, 5, 0, 0, 0, SessionObj.Project.ProjectID);
        DateTime actDate = DateTime.Now.Date;
        MemoryStream outStream = new MemoryStream();
        tempxls.Save(outStream, TFileFormats.Xlsx);

        SessionObj.ExportConfig.outStream = outStream;

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDIngressExport, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    public void doIngressExport()
    {
        SqlDB dataReader;
        
        TIngressMapping mapping = new TIngressMapping(SessionObj.Project.ProjectID);

        ArrayList ingressTargetGroups = new ArrayList();
        
        ArrayList targetGroups = new ArrayList();
        string TIDList = "";
        TIngress webService = new TIngress(SessionObj.Project.ProjectID);

        // Schleife über alle Mitarbeiter (ohne PaperPencil)
        SessionObj.ProcessMessage = "Teilnehmer übertragen ... ";
        int actRow = 2;
        int counter = 0;
        string[] TIDs;
        Dictionary<string, string> ingressTargetGroupsTIDs = new Dictionary<string, string>();
        string XMLData = "";
        string targetGroup;
        dataReader = new SqlDB("Select employeeID FROM orgmanager_employee WHERE (surveytype='0' OR surveytype='2' OR surveytype='3')", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            counter++;

            targetGroup = "";
            XMLData = XMLData + webService.buildXML(dataReader.getInt32(0), ref targetGroup, SessionObj.ingressTargetGroup, SessionObj.indirectIngressTargetGroup, mapping.mapping, SessionObj.AccessCode, SessionObj.Project.ProjectID);
            targetGroups.Add(targetGroup);

            if (counter == 100)
            {
                TProcessStatus.updateProcessInfo(SessionObj.ProcessIDIngressExport, 4, actRow-1, 0, 0, SessionObj.Project.ProjectID);

                XMLData = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><participant_import>" + XMLData + "</participant_import>";

                ingressTargetGroupsTIDs = new Dictionary<string, string>();
                // Daten übertragen
                TIDList = webService.updateParticipant(XMLData);

                // Teilnehmer den Zielgruppen hinzufügen
                counter = 0;
                TIDs = TIDList.Split(',');
                foreach (string aTID in TIDs)
                {
                    string valueList = "";
                    if (ingressTargetGroupsTIDs.TryGetValue(targetGroups[counter].ToString(), out valueList))
                        ingressTargetGroupsTIDs[targetGroups[counter].ToString()] = valueList + "," + aTID;
                    else
                        ingressTargetGroupsTIDs[targetGroups[counter].ToString()] = aTID;

                    counter++;
                }

                // Teilnehmer in Zielgruppen speichern
                foreach (System.Collections.Generic.KeyValuePair<string,string> actTargetGroup in ingressTargetGroupsTIDs)
                {
                    XMLData = "";
                    TIDs = actTargetGroup.Value.Split(',');
                    foreach (string aTID in TIDs)
                    {
                        XMLData = XMLData + "<series><TID>" + aTID + "</TID></series>";
                    }
                    XMLData = "<?xml version =\"1.0\" encoding=\"UTF-8\"?><participant_import>" + XMLData + "</participant_import>";
                    webService.updateParticipantToGroup(XMLData, actTargetGroup.Key);

                    // Zielgruppe merken für Neuzuordnung zur Befragung
                    if (!(ingressTargetGroups.IndexOf(actTargetGroup.Key) >= 0))
                        ingressTargetGroups.Add(actTargetGroup.Key);
                }

                // Rücksetzung
                XMLData = "";
                counter = 0;
                targetGroups.Clear();
            }
            actRow++;
        }
        dataReader.close();
        if (counter !=0 )
        {
            XMLData = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><participant_import>" + XMLData + "</participant_import>";
            // letzte Daten übertragen
            TIDList = webService.updateParticipant(XMLData);

            // Teilnehmer den Zielgruppen hinzufügen
            counter = 0;
            TIDs = TIDList.Split(',');
            foreach (string aTID in TIDs)
            {
                string valueList = "";
                if (ingressTargetGroupsTIDs.TryGetValue(targetGroups[counter].ToString(), out valueList))
                    ingressTargetGroupsTIDs[targetGroups[counter].ToString()] = valueList + "," + aTID;
                else
                    ingressTargetGroupsTIDs[targetGroups[counter].ToString()] = aTID;

                counter++;
            }

            // Teilnehmer in Zielgruppen speichern
            foreach (System.Collections.Generic.KeyValuePair<string, string> actTargetGroup in ingressTargetGroupsTIDs)
            {
                XMLData = "";
                TIDs = actTargetGroup.Value.Split(',');
                XMLData = "";
                foreach (string aTID in TIDs)
                {
                    XMLData = XMLData + "<series><TID>" + aTID + "</TID></series>";
                }
                XMLData = "<?xml version =\"1.0\" encoding=\"UTF-8\"?><participant_import>" + XMLData + "</participant_import>";
                webService.updateParticipantToGroup(XMLData, actTargetGroup.Key);

                // Zielgruppe merken für Neuzuordnung zur Befragung
                if (!(ingressTargetGroups.IndexOf(actTargetGroup.Key) >= 0))
                    ingressTargetGroups.Add(actTargetGroup.Key);
            }
        }

        SessionObj.ProcessMessage = "Gruppen der Befragung neu hinzufügen ... ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDIngressExport, 6, 0, 0, 0, SessionObj.Project.ProjectID);
        DateTime actDate = DateTime.Now.Date;

        // Gruppen der Befragung (neu) zuordnen
        for (int i = 0; i < ingressTargetGroups.Count; i++)
        {
            webService.addGroupToSurvey(ingressTargetGroups[i].ToString());
        }

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDIngressExport, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    #endregion IngressExport

    //-------------------- Threads für User-Export -------------------------------
    #region User-Export
    protected int getColNumber2(string aColName, int languageCount, ExcelFile aXlsFile)
    {
        int colNumber = -1;
        for (int Index = 1; Index <= languageCount; Index++)
        {
            string cellValue = aXlsFile.GetCellValue(1, Index + 2).ToString();
            if (aXlsFile.GetCellValue(1, Index + 2).ToString() == aColName)
                colNumber = Index + 2;
        }
        return colNumber;
    }
    protected int getInt(bool aValue)
    {
        if (aValue)
            return 1;
        else
            return 0;
    }
    public void doUserExport()
    {
        ExcelFile tempxls = new XlsFile(true);
        tempxls.NewFile(40, TExcelFileFormat.v2010);


        # region LoginUserAndMapUser
        // LoginUser zusammenstellen
        System.IO.StringWriter tw = new System.IO.StringWriter();
        tempxls.ActiveSheet = 1;
        tempxls.SheetName = "LoginUser";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "Displayname");
        tempxls.SetCellValue(1, 3, "Password");
        tempxls.SetCellValue(1, 4, "Email");
        tempxls.SetCellValue(1, 5, "Title");
        tempxls.SetCellValue(1, 6, "Name");
        tempxls.SetCellValue(1, 7, "Firstname");
        tempxls.SetCellValue(1, 8, "Street");
        tempxls.SetCellValue(1, 9, "Zipcode");
        tempxls.SetCellValue(1, 10, "City");
        tempxls.SetCellValue(1, 11, "State");
        tempxls.SetCellValue(1, 12, "Language");
        tempxls.SetCellValue(1, 13, "Comment");
        tempxls.SetCellValue(1, 14, "InternalUser");
        tempxls.SetCellValue(1, 15, "GroupID");
        tempxls.SetCellValue(1, 16, "ForceChange");
        tempxls.SetCellValue(1, 17, "Active");
        tempxls.SetCellValue(1, 18, "OrgManager Available");
        tempxls.SetCellValue(1, 19, "Response Available");
        tempxls.SetCellValue(1, 20, "Download Available");
        tempxls.SetCellValue(1, 21, "TeamSpace Available");
        tempxls.SetCellValue(1, 22, "FollowUp Available");
        tempxls.SetCellValue(1, 23, "AccountManager Available");
        tempxls.SetCellValue(1, 24, "OnlineReporting Available");
        tempxls.SetCellValue(1, 25, "Delegate");

        tempxls.ActiveSheet = 2;
        tempxls.SheetName = "UserOrgID";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "O.EditLevel");
        tempxls.SetCellValue(1, 6, "O.Delegate");
        tempxls.SetCellValue(1, 7, "O.Lock");
        tempxls.SetCellValue(1, 8, "O.AlloeEmployeeRead");
        tempxls.SetCellValue(1, 9, "O.AlloeEmployeeEdit");
        tempxls.SetCellValue(1, 10, "O.AllowEmployeeImport");
        tempxls.SetCellValue(1, 11, "O.AllowEmployeeExport");
        tempxls.SetCellValue(1, 12, "O.AllowUnitAdd");
        tempxls.SetCellValue(1, 13, "O.AllowUnitMove");
        tempxls.SetCellValue(1, 14, "O.AllowUnitDelete");
        tempxls.SetCellValue(1, 15, "O.Org2AllowUnitProperty");
        tempxls.SetCellValue(1, 16, "O.Org2AllowReportRecipient");
        tempxls.SetCellValue(1, 17, "O.AllowStructureImport");
        tempxls.SetCellValue(1, 18, "O.AllowStructureExport");
        tempxls.SetCellValue(1, 19, "O.AllowBouncesRead");
        tempxls.SetCellValue(1, 20, "O.AllowBouncesEdit");
        tempxls.SetCellValue(1, 21, "O.AllowBouncesDelete");
        tempxls.SetCellValue(1, 22, "O.AllowBouncesExport");
        tempxls.SetCellValue(1, 23, "O.AllowBouncesImport");
        tempxls.SetCellValue(1, 24, "O.AllowReInvitation");
        tempxls.SetCellValue(1, 25, "O.AllowPostNomination");
        tempxls.SetCellValue(1, 26, "O.AllowPostNominationImport");
        tempxls.SetCellValue(1, 27, "F.Profile");
        tempxls.SetCellValue(1, 28, "F.ReadLevel");
        tempxls.SetCellValue(1, 29, "F.EditLevel");
        tempxls.SetCellValue(1, 30, "F.SumLevel");
        tempxls.SetCellValue(1, 31, "F.AllowCommunications");
        tempxls.SetCellValue(1, 32, "F.AllowMeasures");
        tempxls.SetCellValue(1, 33, "F.Delete");
        tempxls.SetCellValue(1, 34, "F.ExcelSummary");
        tempxls.SetCellValue(1, 35, "F.ReminderOnMeasure");
        tempxls.SetCellValue(1, 36, "F.ReminderOnUnit");
        tempxls.SetCellValue(1, 37, "F.AllowUpUnitMeasures");
        tempxls.SetCellValue(1, 38, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(1, 39, "F.AllowThemes");
        tempxls.SetCellValue(1, 40, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(1, 41, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(1, 42, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(1, 43, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(1, 44, "F.AllowOrders");
        tempxls.SetCellValue(1, 45, "F.AllowReminderOnOrderUnit");
        tempxls.SetCellValue(1, 46, "R.Profile");
        tempxls.SetCellValue(1, 47, "R.ReadLevel");
        tempxls.SetCellValue(1, 48, "D.Profile");
        tempxls.SetCellValue(1, 49, "D.ReadLevel");
        tempxls.SetCellValue(1, 50, "D.Mail");
        tempxls.SetCellValue(1, 51, "D.Download");
        tempxls.SetCellValue(1, 52, "D.Log");
        tempxls.SetCellValue(1, 53, "D.AllowType1");
        tempxls.SetCellValue(1, 54, "D.AllowType2");
        tempxls.SetCellValue(1, 55, "D.AllowType3");
        tempxls.SetCellValue(1, 56, "D.AllowType4");
        tempxls.SetCellValue(1, 57, "D.AllowType5");
        tempxls.SetCellValue(1, 58, "A.EditLevelWorkshop");
        tempxls.SetCellValue(1, 59, "A.ReadLevelStructure");

        tempxls.ActiveSheet = 3;
        tempxls.SheetName = "UserOrgID1";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "O.EditLevel");
        tempxls.SetCellValue(1, 6, "O.Delegate");
        tempxls.SetCellValue(1, 7, "O.Lock");
        tempxls.SetCellValue(1, 8, "O.AlloeEmployeeRead");
        tempxls.SetCellValue(1, 9, "O.AlloeEmployeeEdit");
        tempxls.SetCellValue(1, 10, "O.AllowEmployeeImport");
        tempxls.SetCellValue(1, 11, "O.AllowEmployeeExport");
        tempxls.SetCellValue(1, 12, "O.AllowUnitAdd");
        tempxls.SetCellValue(1, 13, "O.AllowUnitMove");
        tempxls.SetCellValue(1, 14, "O.AllowUnitDelete");
        tempxls.SetCellValue(1, 15, "O.Org2AllowUnitProperty");
        tempxls.SetCellValue(1, 16, "O.Org2AllowReportRecipient");
        tempxls.SetCellValue(1, 17, "O.AllowStructureImport");
        tempxls.SetCellValue(1, 18, "O.AllowStructureExport");
        tempxls.SetCellValue(1, 19, "F.Profile");
        tempxls.SetCellValue(1, 20, "F.ReadLevel");
        tempxls.SetCellValue(1, 21, "F.EditLevel");
        tempxls.SetCellValue(1, 22, "F.SumLevel");
        tempxls.SetCellValue(1, 23, "F.AllowCommunications");
        tempxls.SetCellValue(1, 24, "F.AllowMeasures");
        tempxls.SetCellValue(1, 25, "F.Delete");
        tempxls.SetCellValue(1, 26, "F.ExcelSummary");
        tempxls.SetCellValue(1, 27, "F.ReminderOnMeasure");
        tempxls.SetCellValue(1, 28, "F.ReminderOnUnit");
        tempxls.SetCellValue(1, 29, "F.AllowUpUnitMeasures");
        tempxls.SetCellValue(1, 30, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(1, 31, "F.AllowThemes");
        tempxls.SetCellValue(1, 32, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(1, 33, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(1, 34, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(1, 35, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(1, 36, "F.AllowOrders");
        tempxls.SetCellValue(1, 37, "F.AllowReminderOnOrderUnit");
        tempxls.SetCellValue(1, 38, "R.Profile");
        tempxls.SetCellValue(1, 39, "R.ReadLevel");
        tempxls.SetCellValue(1, 40, "D.Profile");
        tempxls.SetCellValue(1, 41, "D.ReadLevel");
        tempxls.SetCellValue(1, 42, "D.Mail");
        tempxls.SetCellValue(1, 43, "D.Download");
        tempxls.SetCellValue(1, 44, "D.Log");
        tempxls.SetCellValue(1, 45, "D.AllowType1");
        tempxls.SetCellValue(1, 46, "D.AllowType2");
        tempxls.SetCellValue(1, 47, "D.AllowType3");
        tempxls.SetCellValue(1, 48, "D.AllowType4");
        tempxls.SetCellValue(1, 49, "D.AllowType5");
        tempxls.SetCellValue(1, 50, "A.EditLevelWorkshop");
        tempxls.SetCellValue(1, 51, "A.ReadLevelStructure");

        tempxls.ActiveSheet = 4;
        tempxls.SheetName = "UserOrgID2";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "O.EditLevel");
        tempxls.SetCellValue(1, 6, "O.Delegate");
        tempxls.SetCellValue(1, 7, "O.Lock");
        tempxls.SetCellValue(1, 8, "O.AlloeEmployeeRead");
        tempxls.SetCellValue(1, 9, "O.AlloeEmployeeEdit");
        tempxls.SetCellValue(1, 10, "O.AllowEmployeeImport");
        tempxls.SetCellValue(1, 11, "O.AllowEmployeeExport");
        tempxls.SetCellValue(1, 12, "O.AllowUnitAdd");
        tempxls.SetCellValue(1, 13, "O.AllowUnitMove");
        tempxls.SetCellValue(1, 14, "O.AllowUnitDelete");
        tempxls.SetCellValue(1, 15, "O.Org2AllowUnitProperty");
        tempxls.SetCellValue(1, 16, "O.Org2AllowReportRecipient");
        tempxls.SetCellValue(1, 17, "O.AllowStructureImport");
        tempxls.SetCellValue(1, 18, "O.AllowStructureExport");
        tempxls.SetCellValue(1, 19, "F.Profile");
        tempxls.SetCellValue(1, 20, "F.ReadLevel");
        tempxls.SetCellValue(1, 21, "F.EditLevel");
        tempxls.SetCellValue(1, 22, "F.SumLevel");
        tempxls.SetCellValue(1, 23, "F.AllowCommunications");
        tempxls.SetCellValue(1, 24, "F.AllowMeasures");
        tempxls.SetCellValue(1, 25, "F.Delete");
        tempxls.SetCellValue(1, 26, "F.ExcelSummary");
        tempxls.SetCellValue(1, 27, "F.ReminderOnMeasure");
        tempxls.SetCellValue(1, 28, "F.ReminderOnUnit");
        tempxls.SetCellValue(1, 29, "F.AllowUpUnitMeasures");
        tempxls.SetCellValue(1, 30, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(1, 31, "F.AllowThemes");
        tempxls.SetCellValue(1, 32, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(1, 33, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(1, 34, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(1, 35, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(1, 36, "F.AllowOrders");
        tempxls.SetCellValue(1, 37, "F.AllowReminderOnOrderUnit");
        tempxls.SetCellValue(1, 38, "R.Profile");
        tempxls.SetCellValue(1, 39, "R.ReadLevel");
        tempxls.SetCellValue(1, 40, "D.Profile");
        tempxls.SetCellValue(1, 41, "D.ReadLevel");
        tempxls.SetCellValue(1, 42, "D.Mail");
        tempxls.SetCellValue(1, 43, "D.Download");
        tempxls.SetCellValue(1, 44, "D.Log");
        tempxls.SetCellValue(1, 45, "D.AllowType1");
        tempxls.SetCellValue(1, 46, "D.AllowType2");
        tempxls.SetCellValue(1, 47, "D.AllowType3");
        tempxls.SetCellValue(1, 48, "D.AllowType4");
        tempxls.SetCellValue(1, 49, "D.AllowType5");
        tempxls.SetCellValue(1, 50, "A.EditLevelWorkshop");
        tempxls.SetCellValue(1, 51, "A.ReadLevelStructure");

        tempxls.ActiveSheet = 5;
        tempxls.SheetName = "UserOrgID3";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "O.EditLevel");
        tempxls.SetCellValue(1, 6, "O.Delegate");
        tempxls.SetCellValue(1, 7, "O.Lock");
        tempxls.SetCellValue(1, 8, "O.AlloeEmployeeRead");
        tempxls.SetCellValue(1, 9, "O.AlloeEmployeeEdit");
        tempxls.SetCellValue(1, 10, "O.AllowEmployeeImport");
        tempxls.SetCellValue(1, 11, "O.AllowEmployeeExport");
        tempxls.SetCellValue(1, 12, "O.AllowUnitAdd");
        tempxls.SetCellValue(1, 13, "O.AllowUnitMove");
        tempxls.SetCellValue(1, 14, "O.AllowUnitDelete");
        tempxls.SetCellValue(1, 15, "O.Org2AllowUnitProperty");
        tempxls.SetCellValue(1, 16, "O.Org2AllowReportRecipient");
        tempxls.SetCellValue(1, 17, "O.AllowStructureImport");
        tempxls.SetCellValue(1, 18, "O.AllowStructureExport");
        tempxls.SetCellValue(1, 19, "F.Profile");
        tempxls.SetCellValue(1, 20, "F.ReadLevel");
        tempxls.SetCellValue(1, 21, "F.EditLevel");
        tempxls.SetCellValue(1, 22, "F.SumLevel");
        tempxls.SetCellValue(1, 23, "F.AllowCommunications");
        tempxls.SetCellValue(1, 24, "F.AllowMeasures");
        tempxls.SetCellValue(1, 25, "F.Delete");
        tempxls.SetCellValue(1, 26, "F.ExcelSummary");
        tempxls.SetCellValue(1, 27, "F.ReminderOnMeasure");
        tempxls.SetCellValue(1, 28, "F.ReminderOnUnit");
        tempxls.SetCellValue(1, 29, "F.AllowUpUnitMeasures");
        tempxls.SetCellValue(1, 30, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(1, 31, "F.AllowThemes");
        tempxls.SetCellValue(1, 32, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(1, 33, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(1, 34, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(1, 35, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(1, 36, "F.AllowOrders");
        tempxls.SetCellValue(1, 37, "F.AllowReminderOnOrderUnit");
        tempxls.SetCellValue(1, 38, "R.Profile");
        tempxls.SetCellValue(1, 39, "R.ReadLevel");
        tempxls.SetCellValue(1, 40, "D.Profile");
        tempxls.SetCellValue(1, 41, "D.ReadLevel");
        tempxls.SetCellValue(1, 42, "D.Mail");
        tempxls.SetCellValue(1, 43, "D.Download");
        tempxls.SetCellValue(1, 44, "D.Log");
        tempxls.SetCellValue(1, 45, "D.AllowType1");
        tempxls.SetCellValue(1, 46, "D.AllowType2");
        tempxls.SetCellValue(1, 47, "D.AllowType3");
        tempxls.SetCellValue(1, 48, "D.AllowType4");
        tempxls.SetCellValue(1, 49, "D.AllowType5");
        tempxls.SetCellValue(1, 50, "A.EditLevelWorkshop");
        tempxls.SetCellValue(1, 51, "A.ReadLevelStructure");

        tempxls.ActiveSheet = 6;
        tempxls.SheetName = "UserBack";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "A.ReadLevel");

        tempxls.ActiveSheet = 7;
        tempxls.SheetName = "UserBack1";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "A.ReadLevel");

        tempxls.ActiveSheet = 8;
        tempxls.SheetName = "UserBack2";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "A.ReadLevel");

        tempxls.ActiveSheet = 9;
        tempxls.SheetName = "UserBack3";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "OrgID");
        tempxls.SetCellValue(1, 3, "O.Profile");
        tempxls.SetCellValue(1, 4, "O.ReadLevel");
        tempxls.SetCellValue(1, 5, "A.ReadLevel");

        tempxls.ActiveSheet = 10;
        tempxls.SheetName = "UserContainer";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "ContainerID");
        tempxls.SetCellValue(1, 3, "T.Containerprofile");
        tempxls.SetCellValue(1, 4, "T.AllowUpload");
        tempxls.SetCellValue(1, 5, "T.AllowDownload");
        tempxls.SetCellValue(1, 6, "T.AllowDeleteOwnFiles");
        tempxls.SetCellValue(1, 7, "T.AllowDeleteAllFiles");
        tempxls.SetCellValue(1, 8, "T.AllowAccessOwnFilesWithoutPassword");
        tempxls.SetCellValue(1, 9, "T.AllowAccessAllFilesWithoutPassword");
        tempxls.SetCellValue(1, 10, "T.AllowCreateFolder");
        tempxls.SetCellValue(1, 11, "T.AllowDeleteOwnFolder");
        tempxls.SetCellValue(1, 12, "T.AllowDeleteAllFolder");
        tempxls.SetCellValue(1, 13, "T.AllowResetPassword");
        tempxls.SetCellValue(1, 14, "T.AllowTakeOwnership");

        tempxls.ActiveSheet = 11;
        tempxls.SheetName = "UserRights";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "T.Taskprofile");
        tempxls.SetCellValue(1, 3, "T.AllowTaskRead");
        tempxls.SetCellValue(1, 4, "T.AllowTaskCreateEditDelete");
        tempxls.SetCellValue(1, 5, "T.AllowTaskFileUpload");
        tempxls.SetCellValue(1, 6, "T.AllowTaskExport");
        tempxls.SetCellValue(1, 7, "T.AllowTaskImport");
        tempxls.SetCellValue(1, 8, "T.Contactprofile");
        tempxls.SetCellValue(1, 9, "T.AllowContactRead");
        tempxls.SetCellValue(1, 10, "T.AllowContactCreateEditDelete");
        tempxls.SetCellValue(1, 11, "T.AllowContactExport");
        tempxls.SetCellValue(1, 12, "T.AllowContactImport");
        tempxls.SetCellValue(1, 13, "T.A.AllowAccountTransfer");
        tempxls.SetCellValue(1, 14, "T.A.AllowAccountManagement");
        tempxls.SetCellValue(1, 15, "T.Translateprofile");
        tempxls.SetCellValue(1, 16, "T.AllowTranslateRead");
        tempxls.SetCellValue(1, 17, "T.AllowTranslateEdit");
        tempxls.SetCellValue(1, 18, "T.AllowTranslateRelease");
        tempxls.SetCellValue(1, 19, "T.AllowTranslateNew");
        tempxls.SetCellValue(1, 20, "T.AllowTranslateDelete");

        tempxls.ActiveSheet = 12;
        tempxls.SheetName = "UserTaskRole";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "T.RoleID");

        tempxls.ActiveSheet = 13;
        tempxls.SheetName = "UserGroups";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "GroupID");

        tempxls.ActiveSheet = 14;
        tempxls.SheetName = "UserGroup";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "A.GroupID");

        tempxls.ActiveSheet = 15;
        tempxls.SheetName = "UserTranslateLanguage";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "UserID");
        tempxls.SetCellValue(1, 2, "T.TranslateLanguageID");

        // LoginUser ausgeben
        int actRow = 2;
        int actRow2 = 2;
        int actRow3 = 2;
        int actRow4 = 2;
        int actRow5 = 2;
        int actRow6 = 2;
        int actRow7 = 2;
        int actRow8 = 2;
        int actRow9 = 2;
        int actRow10 = 2;
        int actRow11 = 2;
        int actRow12 = 2;
        int actRow14 = 2;
        int actRow15 = 2;
        TUser tempUser;
        SessionObj.ProcessMessage = "LoginUser ausgeben ... ";
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT userID FROM loginuser", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            TProcessStatus.updateProcessInfo(SessionObj.ProcessIDUserExport, 3, actRow, 0, 0, SessionObj.Project.ProjectID);
            tempUser = new TUser(dataReader.getString(0), SessionObj.Project.ProjectID);

            tempxls.ActiveSheet = 1;
            tempxls.SetCellValue(actRow, 1, tempUser.UserID);
            tempxls.SetCellValue(actRow, 2, tempUser.DisplayName);
            tempxls.SetCellValue(actRow, 3, "");
            tempxls.SetCellValue(actRow, 4, tempUser.Email);
            tempxls.SetCellValue(actRow, 5, tempUser.Title);
            tempxls.SetCellValue(actRow, 6, tempUser.Name);
            tempxls.SetCellValue(actRow, 7, tempUser.Firstname);
            tempxls.SetCellValue(actRow, 8, tempUser.Street);
            tempxls.SetCellValue(actRow, 9, tempUser.Zipcode);
            tempxls.SetCellValue(actRow, 10, tempUser.City);
            tempxls.SetCellValue(actRow, 11, tempUser.State);
            tempxls.SetCellValue(actRow, 12, tempUser.Language);
            tempxls.SetCellValue(actRow, 13, tempUser.Comment);
            tempxls.SetCellValue(actRow, 14, getInt(tempUser.IsInternalUser));
            tempxls.SetCellValue(actRow, 15, tempUser.GroupID);
            tempxls.SetCellValue(actRow, 16, getInt(tempUser.ForceChange));
            tempxls.SetCellValue(actRow, 17, getInt(tempUser.Active));
            tempxls.SetCellValue(actRow, 18, getInt(tempUser.getToolAvailable("orgmanager")));
            tempxls.SetCellValue(actRow, 19, getInt(tempUser.getToolAvailable("response")));
            tempxls.SetCellValue(actRow, 20, getInt(tempUser.getToolAvailable("download")));
            tempxls.SetCellValue(actRow, 21, getInt(tempUser.getToolAvailable("teamspace")));
            tempxls.SetCellValue(actRow, 22, getInt(tempUser.getToolAvailable("followup")));
            tempxls.SetCellValue(actRow, 23, getInt(tempUser.getToolAvailable("accountmanager")));
            tempxls.SetCellValue(actRow, 24, getInt(tempUser.getToolAvailable("onlinereporting")));
            tempxls.SetCellValue(actRow, 25, getInt(tempUser.IsDelegate));
            actRow++;

            tempxls.ActiveSheet = 2;
            foreach (TUser.TOrgRight tempRight in tempUser.orgRights)
            {
                tempxls.SetCellValue(actRow2, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow2, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow2, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow2, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow2, 5, tempRight.OEditLevel);
                tempxls.SetCellValue(actRow2, 6, getInt(tempRight.OAllowDelegate));
                tempxls.SetCellValue(actRow2, 7, getInt(tempRight.OAllowLock));
                tempxls.SetCellValue(actRow2, 8, getInt(tempRight.OAllowEmployeeRead));
                tempxls.SetCellValue(actRow2, 9, getInt(tempRight.OAllowEmployeeEdit));
                tempxls.SetCellValue(actRow2, 10, getInt(tempRight.OAllowEmployeeImport));
                tempxls.SetCellValue(actRow2, 11, getInt(tempRight.OAllowEmployeeExport));
                tempxls.SetCellValue(actRow2, 12, getInt(tempRight.OAllowUnitAdd));
                tempxls.SetCellValue(actRow2, 13, getInt(tempRight.OAllowUnitMove));
                tempxls.SetCellValue(actRow2, 14, getInt(tempRight.OAllowUnitDelete));
                tempxls.SetCellValue(actRow2, 15, getInt(tempRight.OAllowUnitProperty));
                tempxls.SetCellValue(actRow2, 16, getInt(tempRight.OAllowReportRecipient));
                tempxls.SetCellValue(actRow2, 17, getInt(tempRight.OAllowStructureImport));
                tempxls.SetCellValue(actRow2, 18, getInt(tempRight.OAllowStructureExport));
                tempxls.SetCellValue(actRow2, 19, getInt(tempRight.OAllowBouncesRead));
                tempxls.SetCellValue(actRow2, 20, getInt(tempRight.OAllowBouncesEdit));
                tempxls.SetCellValue(actRow2, 21, getInt(tempRight.OAllowBouncesDelete));
                tempxls.SetCellValue(actRow2, 22, getInt(tempRight.OAllowBouncesExport));
                tempxls.SetCellValue(actRow2, 23, getInt(tempRight.OAllowBouncesImport));
                tempxls.SetCellValue(actRow2, 24, getInt(tempRight.OAllowReInvitation));
                tempxls.SetCellValue(actRow2, 25, getInt(tempRight.OAllowPostNomination));
                tempxls.SetCellValue(actRow2, 26, getInt(tempRight.OAllowPostNominationImport));
                tempxls.SetCellValue(actRow2, 27, tempRight.FProfile);
                tempxls.SetCellValue(actRow2, 28, tempRight.FReadLevel);
                tempxls.SetCellValue(actRow2, 29, tempRight.FEditLevel);
                tempxls.SetCellValue(actRow2, 30, tempRight.FSumLevel);
                tempxls.SetCellValue(actRow2, 31, getInt(tempRight.FAllowCommunications));
                tempxls.SetCellValue(actRow2, 32, getInt(tempRight.FAllowMeasures));
                tempxls.SetCellValue(actRow2, 33, getInt(tempRight.FAllowDelete));
                tempxls.SetCellValue(actRow2, 34, getInt(tempRight.FAllowExcelExport));
                tempxls.SetCellValue(actRow2, 35, getInt(tempRight.FAllowReminderOnMeasure));
                tempxls.SetCellValue(actRow2, 36, getInt(tempRight.FAllowReminderOnUnit));
                tempxls.SetCellValue(actRow2, 37, getInt(tempRight.FAllowUpUnitsMeasures));
                tempxls.SetCellValue(actRow2, 38, getInt(tempRight.FAllowColleaguesMeasures));
                tempxls.SetCellValue(actRow2, 39, getInt(tempRight.FAllowThemes));
                tempxls.SetCellValue(actRow2, 40, getInt(tempRight.FAllowReminderOnThemes));
                tempxls.SetCellValue(actRow2, 41, getInt(tempRight.FAllowReminderOnThemeUnit));
                tempxls.SetCellValue(actRow2, 42, getInt(tempRight.FAllowTopUnitsThemes));
                tempxls.SetCellValue(actRow2, 43, getInt(tempRight.FAllowColleaguesThemes));
                tempxls.SetCellValue(actRow2, 44, getInt(tempRight.FAllowOrder));
                tempxls.SetCellValue(actRow2, 45, getInt(tempRight.FAllowReminderOnOrderUnit));
                tempxls.SetCellValue(actRow2, 46, tempRight.RProfile);
                tempxls.SetCellValue(actRow2, 47, tempRight.RReadLevel);
                tempxls.SetCellValue(actRow2, 48, tempRight.DProfile);
                tempxls.SetCellValue(actRow2, 49, tempRight.DReadLevel);
                tempxls.SetCellValue(actRow2, 50, getInt(tempRight.DAllowMail));
                tempxls.SetCellValue(actRow2, 51, getInt(tempRight.DAllowDownload));
                tempxls.SetCellValue(actRow2, 52, getInt(tempRight.DAllowLog));
                tempxls.SetCellValue(actRow2, 53, getInt(tempRight.DAllowType1));
                tempxls.SetCellValue(actRow2, 54, getInt(tempRight.DAllowType2));
                tempxls.SetCellValue(actRow2, 55, getInt(tempRight.DAllowType3));
                tempxls.SetCellValue(actRow2, 56, getInt(tempRight.DAllowType4));
                tempxls.SetCellValue(actRow2, 57, getInt(tempRight.DAllowType5));
                tempxls.SetCellValue(actRow2, 58, tempRight.AEditLevelWorkshop);
                tempxls.SetCellValue(actRow2, 59, tempRight.AReadLevelStructure);
                actRow2++;
            }

            tempxls.ActiveSheet = 3;
            foreach (TUser.TOrg1Right tempRight in tempUser.org1Rights)
            {
                tempxls.SetCellValue(actRow3, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow3, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow3, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow3, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow3, 5, tempRight.OEditLevel);
                tempxls.SetCellValue(actRow3, 6, getInt(tempRight.OAllowDelegate));
                tempxls.SetCellValue(actRow3, 7, getInt(tempRight.OAllowLock));
                tempxls.SetCellValue(actRow3, 8, getInt(tempRight.OAllowEmployeeRead));
                tempxls.SetCellValue(actRow3, 9, getInt(tempRight.OAllowEmployeeEdit));
                tempxls.SetCellValue(actRow3, 10, getInt(tempRight.OAllowEmployeeImport));
                tempxls.SetCellValue(actRow3, 11, getInt(tempRight.OAllowEmployeeExport));
                tempxls.SetCellValue(actRow3, 12, getInt(tempRight.OAllowUnitAdd));
                tempxls.SetCellValue(actRow3, 13, getInt(tempRight.OAllowUnitMove));
                tempxls.SetCellValue(actRow3, 14, getInt(tempRight.OAllowUnitDelete));
                tempxls.SetCellValue(actRow3, 15, getInt(tempRight.OAllowUnitProperty));
                tempxls.SetCellValue(actRow3, 16, getInt(tempRight.OAllowReportRecipient));
                tempxls.SetCellValue(actRow3, 17, getInt(tempRight.OAllowStructureImport));
                tempxls.SetCellValue(actRow3, 18, getInt(tempRight.OAllowStructureExport));
                tempxls.SetCellValue(actRow3, 19, tempRight.FProfile);
                tempxls.SetCellValue(actRow3, 20, tempRight.FReadLevel);
                tempxls.SetCellValue(actRow3, 21, tempRight.FEditLevel);
                tempxls.SetCellValue(actRow3, 22, tempRight.FSumLevel);
                tempxls.SetCellValue(actRow3, 23, getInt(tempRight.FAllowCommunications));
                tempxls.SetCellValue(actRow3, 24, getInt(tempRight.FAllowMeasures));
                tempxls.SetCellValue(actRow3, 25, getInt(tempRight.FAllowDelete));
                tempxls.SetCellValue(actRow3, 26, getInt(tempRight.FAllowExcelExport));
                tempxls.SetCellValue(actRow3, 27, getInt(tempRight.FAllowReminderOnMeasure));
                tempxls.SetCellValue(actRow3, 28, getInt(tempRight.FAllowReminderOnUnit));
                tempxls.SetCellValue(actRow3, 29, getInt(tempRight.FAllowUpUnitsMeasures));
                tempxls.SetCellValue(actRow3, 30, getInt(tempRight.FAllowColleaguesMeasures));
                tempxls.SetCellValue(actRow3, 31, getInt(tempRight.FAllowThemes));
                tempxls.SetCellValue(actRow3, 32, getInt(tempRight.FAllowReminderOnThemes));
                tempxls.SetCellValue(actRow3, 33, getInt(tempRight.FAllowReminderOnThemeUnit));
                tempxls.SetCellValue(actRow3, 34, getInt(tempRight.FAllowTopUnitsThemes));
                tempxls.SetCellValue(actRow3, 35, getInt(tempRight.FAllowColleaguesThemes));
                tempxls.SetCellValue(actRow3, 36, getInt(tempRight.FAllowOrder));
                tempxls.SetCellValue(actRow3, 37, getInt(tempRight.FAllowReminderOnOrderUnit));
                tempxls.SetCellValue(actRow3, 38, tempRight.RProfile);
                tempxls.SetCellValue(actRow3, 39, tempRight.RReadLevel);
                tempxls.SetCellValue(actRow3, 40, tempRight.DProfile);
                tempxls.SetCellValue(actRow3, 41, tempRight.DReadLevel);
                tempxls.SetCellValue(actRow3, 42, getInt(tempRight.DAllowMail));
                tempxls.SetCellValue(actRow3, 43, getInt(tempRight.DAllowDownload));
                tempxls.SetCellValue(actRow3, 44, getInt(tempRight.DAllowLog));
                tempxls.SetCellValue(actRow3, 45, getInt(tempRight.DAllowType1));
                tempxls.SetCellValue(actRow3, 46, getInt(tempRight.DAllowType2));
                tempxls.SetCellValue(actRow3, 47, getInt(tempRight.DAllowType3));
                tempxls.SetCellValue(actRow3, 48, getInt(tempRight.DAllowType4));
                tempxls.SetCellValue(actRow3, 49, getInt(tempRight.DAllowType5));
                tempxls.SetCellValue(actRow3, 50, tempRight.AEditLevelWorkshop);
                tempxls.SetCellValue(actRow3, 51, tempRight.AReadLevelStructure);
                actRow3++;
            }

            tempxls.ActiveSheet = 4;
            foreach (TUser.TOrg2Right tempRight in tempUser.org2Rights)
            {
                tempxls.SetCellValue(actRow4, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow4, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow4, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow4, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow4, 5, tempRight.OEditLevel);
                tempxls.SetCellValue(actRow4, 6, getInt(tempRight.OAllowDelegate));
                tempxls.SetCellValue(actRow4, 7, getInt(tempRight.OAllowLock));
                tempxls.SetCellValue(actRow4, 8, getInt(tempRight.OAllowEmployeeRead));
                tempxls.SetCellValue(actRow4, 9, getInt(tempRight.OAllowEmployeeEdit));
                tempxls.SetCellValue(actRow4, 10, getInt(tempRight.OAllowEmployeeImport));
                tempxls.SetCellValue(actRow4, 11, getInt(tempRight.OAllowEmployeeExport));
                tempxls.SetCellValue(actRow4, 12, getInt(tempRight.OAllowUnitAdd));
                tempxls.SetCellValue(actRow4, 13, getInt(tempRight.OAllowUnitMove));
                tempxls.SetCellValue(actRow4, 14, getInt(tempRight.OAllowUnitDelete));
                tempxls.SetCellValue(actRow4, 15, getInt(tempRight.OAllowUnitProperty));
                tempxls.SetCellValue(actRow4, 16, getInt(tempRight.OAllowReportRecipient));
                tempxls.SetCellValue(actRow4, 17, getInt(tempRight.OAllowStructureImport));
                tempxls.SetCellValue(actRow4, 18, getInt(tempRight.OAllowStructureExport));
                tempxls.SetCellValue(actRow4, 19, tempRight.FProfile);
                tempxls.SetCellValue(actRow4, 20, tempRight.FReadLevel);
                tempxls.SetCellValue(actRow4, 21, tempRight.FEditLevel);
                tempxls.SetCellValue(actRow4, 22, tempRight.FSumLevel);
                tempxls.SetCellValue(actRow4, 23, getInt(tempRight.FAllowCommunications));
                tempxls.SetCellValue(actRow4, 24, getInt(tempRight.FAllowMeasures));
                tempxls.SetCellValue(actRow4, 25, getInt(tempRight.FAllowDelete));
                tempxls.SetCellValue(actRow4, 26, getInt(tempRight.FAllowExcelExport));
                tempxls.SetCellValue(actRow4, 27, getInt(tempRight.FAllowReminderOnMeasure));
                tempxls.SetCellValue(actRow4, 28, getInt(tempRight.FAllowReminderOnUnit));
                tempxls.SetCellValue(actRow4, 29, getInt(tempRight.FAllowUpUnitsMeasures));
                tempxls.SetCellValue(actRow4, 30, getInt(tempRight.FAllowColleaguesMeasures));
                tempxls.SetCellValue(actRow4, 31, getInt(tempRight.FAllowThemes));
                tempxls.SetCellValue(actRow4, 32, getInt(tempRight.FAllowReminderOnThemes));
                tempxls.SetCellValue(actRow4, 33, getInt(tempRight.FAllowReminderOnThemeUnit));
                tempxls.SetCellValue(actRow4, 34, getInt(tempRight.FAllowTopUnitsThemes));
                tempxls.SetCellValue(actRow4, 35, getInt(tempRight.FAllowColleaguesThemes));
                tempxls.SetCellValue(actRow4, 36, getInt(tempRight.FAllowOrder));
                tempxls.SetCellValue(actRow4, 37, getInt(tempRight.FAllowReminderOnOrderUnit));
                tempxls.SetCellValue(actRow4, 38, tempRight.RProfile);
                tempxls.SetCellValue(actRow4, 39, tempRight.RReadLevel);
                tempxls.SetCellValue(actRow4, 40, tempRight.DProfile);
                tempxls.SetCellValue(actRow4, 41, tempRight.DReadLevel);
                tempxls.SetCellValue(actRow4, 42, getInt(tempRight.DAllowMail));
                tempxls.SetCellValue(actRow4, 43, getInt(tempRight.DAllowDownload));
                tempxls.SetCellValue(actRow4, 44, getInt(tempRight.DAllowLog));
                tempxls.SetCellValue(actRow4, 45, getInt(tempRight.DAllowType1));
                tempxls.SetCellValue(actRow4, 46, getInt(tempRight.DAllowType2));
                tempxls.SetCellValue(actRow4, 47, getInt(tempRight.DAllowType3));
                tempxls.SetCellValue(actRow4, 48, getInt(tempRight.DAllowType4));
                tempxls.SetCellValue(actRow4, 49, getInt(tempRight.DAllowType5));
                tempxls.SetCellValue(actRow4, 50, tempRight.AEditLevelWorkshop);
                tempxls.SetCellValue(actRow4, 51, tempRight.AReadLevelStructure);
                actRow4++;
            }

            tempxls.ActiveSheet = 5;
            foreach (TUser.TOrg3Right tempRight in tempUser.org3Rights)
            {
                tempxls.SetCellValue(actRow5, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow5, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow5, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow5, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow5, 5, tempRight.OEditLevel);
                tempxls.SetCellValue(actRow5, 6, getInt(tempRight.OAllowDelegate));
                tempxls.SetCellValue(actRow5, 7, getInt(tempRight.OAllowLock));
                tempxls.SetCellValue(actRow5, 8, getInt(tempRight.OAllowEmployeeRead));
                tempxls.SetCellValue(actRow5, 9, getInt(tempRight.OAllowEmployeeEdit));
                tempxls.SetCellValue(actRow5, 10, getInt(tempRight.OAllowEmployeeImport));
                tempxls.SetCellValue(actRow5, 11, getInt(tempRight.OAllowEmployeeExport));
                tempxls.SetCellValue(actRow5, 12, getInt(tempRight.OAllowUnitAdd));
                tempxls.SetCellValue(actRow5, 13, getInt(tempRight.OAllowUnitMove));
                tempxls.SetCellValue(actRow5, 14, getInt(tempRight.OAllowUnitDelete));
                tempxls.SetCellValue(actRow5, 15, getInt(tempRight.OAllowUnitProperty));
                tempxls.SetCellValue(actRow5, 16, getInt(tempRight.OAllowReportRecipient));
                tempxls.SetCellValue(actRow5, 17, getInt(tempRight.OAllowStructureImport));
                tempxls.SetCellValue(actRow5, 18, getInt(tempRight.OAllowStructureExport));
                tempxls.SetCellValue(actRow5, 19, tempRight.FProfile);
                tempxls.SetCellValue(actRow5, 20, tempRight.FReadLevel);
                tempxls.SetCellValue(actRow5, 21, tempRight.FEditLevel);
                tempxls.SetCellValue(actRow5, 22, tempRight.FSumLevel);
                tempxls.SetCellValue(actRow5, 23, getInt(tempRight.FAllowCommunications));
                tempxls.SetCellValue(actRow5, 24, getInt(tempRight.FAllowMeasures));
                tempxls.SetCellValue(actRow5, 25, getInt(tempRight.FAllowDelete));
                tempxls.SetCellValue(actRow5, 26, getInt(tempRight.FAllowExcelExport));
                tempxls.SetCellValue(actRow5, 27, getInt(tempRight.FAllowReminderOnMeasure));
                tempxls.SetCellValue(actRow5, 28, getInt(tempRight.FAllowReminderOnUnit));
                tempxls.SetCellValue(actRow5, 29, getInt(tempRight.FAllowUpUnitsMeasures));
                tempxls.SetCellValue(actRow5, 30, getInt(tempRight.FAllowColleaguesMeasures));
                tempxls.SetCellValue(actRow5, 31, getInt(tempRight.FAllowThemes));
                tempxls.SetCellValue(actRow5, 32, getInt(tempRight.FAllowReminderOnThemes));
                tempxls.SetCellValue(actRow5, 33, getInt(tempRight.FAllowReminderOnThemeUnit));
                tempxls.SetCellValue(actRow5, 34, getInt(tempRight.FAllowTopUnitsThemes));
                tempxls.SetCellValue(actRow5, 35, getInt(tempRight.FAllowColleaguesThemes));
                tempxls.SetCellValue(actRow5, 36, getInt(tempRight.FAllowOrder));
                tempxls.SetCellValue(actRow5, 37, getInt(tempRight.FAllowReminderOnOrderUnit));
                tempxls.SetCellValue(actRow5, 38, tempRight.RProfile);
                tempxls.SetCellValue(actRow5, 39, tempRight.RReadLevel);
                tempxls.SetCellValue(actRow5, 40, tempRight.DProfile);
                tempxls.SetCellValue(actRow5, 41, tempRight.DReadLevel);
                tempxls.SetCellValue(actRow5, 42, getInt(tempRight.DAllowMail));
                tempxls.SetCellValue(actRow5, 43, getInt(tempRight.DAllowDownload));
                tempxls.SetCellValue(actRow5, 44, getInt(tempRight.DAllowLog));
                tempxls.SetCellValue(actRow5, 45, getInt(tempRight.DAllowType1));
                tempxls.SetCellValue(actRow5, 46, getInt(tempRight.DAllowType2));
                tempxls.SetCellValue(actRow5, 47, getInt(tempRight.DAllowType3));
                tempxls.SetCellValue(actRow5, 48, getInt(tempRight.DAllowType4));
                tempxls.SetCellValue(actRow5, 49, getInt(tempRight.DAllowType5));
                tempxls.SetCellValue(actRow5, 50, tempRight.AEditLevelWorkshop);
                tempxls.SetCellValue(actRow5, 51, tempRight.AReadLevelStructure);
                actRow5++;
            }

            tempxls.ActiveSheet = 6;
            foreach (TUser.TBackRight tempRight in tempUser.backRights)
            {
                tempxls.SetCellValue(actRow6, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow6, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow6, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow6, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow6, 5, tempRight.AReadLevel);
                actRow6++;
            }

            tempxls.ActiveSheet = 7;
            foreach (TUser.TBack1Right tempRight in tempUser.back1Rights)
            {
                tempxls.SetCellValue(actRow7, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow7, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow7, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow7, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow7, 5, tempRight.AReadLevel);
                actRow7++;
            }

            tempxls.ActiveSheet = 8;
            foreach (TUser.TBack2Right tempRight in tempUser.back2Rights)
            {
                tempxls.SetCellValue(actRow8, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow8, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow8, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow8, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow8, 5, tempRight.AReadLevel);
                actRow8++;
            }

            tempxls.ActiveSheet = 9;
            foreach (TUser.TBack3Right tempRight in tempUser.back3Rights)
            {
                tempxls.SetCellValue(actRow9, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow9, 2, tempRight.OrgID);
                tempxls.SetCellValue(actRow9, 3, tempRight.OProfile);
                tempxls.SetCellValue(actRow9, 4, tempRight.OReadLevel);
                tempxls.SetCellValue(actRow9, 5, tempRight.AReadLevel);
                actRow9++;
            }

            tempxls.ActiveSheet = 10;
            foreach (TUser.TContainerRight tempRight in tempUser.containerRights)
            {
                tempxls.SetCellValue(actRow10, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow10, 2, tempRight.ContainerID);
                tempxls.SetCellValue(actRow10, 3, tempRight.TContainerProfile);
                tempxls.SetCellValue(actRow10, 4, getInt(tempRight.TAllowUpload));
                tempxls.SetCellValue(actRow10, 5, getInt(tempRight.TAllowDownload));
                tempxls.SetCellValue(actRow10, 6, getInt(tempRight.TAllowDeleteOwnFiles));
                tempxls.SetCellValue(actRow10, 7, getInt(tempRight.TAllowDeleteAllFiles));
                tempxls.SetCellValue(actRow10, 8, getInt(tempRight.TAllowAccessOwnFilesWithoutPassword));
                tempxls.SetCellValue(actRow10, 9, getInt(tempRight.TAllowAccessAllFilesWithoutPassword));
                tempxls.SetCellValue(actRow10, 10, getInt(tempRight.TAllowCreateFolder));
                tempxls.SetCellValue(actRow10, 11, getInt(tempRight.TAllowDeleteOwnFolder));
                tempxls.SetCellValue(actRow10, 12, getInt(tempRight.TAllowDeleteAllFolder));
                tempxls.SetCellValue(actRow10, 13, getInt(tempRight.TAllowResetPassword));
                tempxls.SetCellValue(actRow10, 14, getInt(tempRight.TAllowTakeOwnership));
                actRow10++;
            }

            tempxls.ActiveSheet = 11;
            tempxls.SetCellValue(actRow11, 1, tempUser.UserID);
            tempxls.SetCellValue(actRow11, 2, tempUser.userRights.TTaskProfile);
            tempxls.SetCellValue(actRow11, 3, getInt(tempUser.userRights.TAllowTaskRead));
            tempxls.SetCellValue(actRow11, 4, getInt(tempUser.userRights.TAllowTaskCreateEditDelete));
            tempxls.SetCellValue(actRow11, 5, getInt(tempUser.userRights.TAllowTaskFileUpload));
            tempxls.SetCellValue(actRow11, 6, getInt(tempUser.userRights.TAllowTaskExport));
            tempxls.SetCellValue(actRow11, 7, getInt(tempUser.userRights.TAllowTaskImport));
            tempxls.SetCellValue(actRow11, 8, getInt(tempUser.userRights.TAllowContactRead));
            tempxls.SetCellValue(actRow11, 9, tempUser.userRights.TContactProfile);
            tempxls.SetCellValue(actRow11, 10, getInt(tempUser.userRights.TAllowContactCreateEditDelete));
            tempxls.SetCellValue(actRow11, 11, getInt(tempUser.userRights.TAllowContactExport));
            tempxls.SetCellValue(actRow11, 12, getInt(tempUser.userRights.TAllowContactImport));
            tempxls.SetCellValue(actRow11, 13, getInt(tempUser.userRights.AAllowAccountTransfer));
            tempxls.SetCellValue(actRow11, 14, getInt(tempUser.userRights.AAllowAccountManagement));
            tempxls.SetCellValue(actRow11, 15, tempUser.userRights.TTranslateProfile);
            tempxls.SetCellValue(actRow11, 16, getInt(tempUser.userRights.TAllowTranslateRead));
            tempxls.SetCellValue(actRow11, 17, getInt(tempUser.userRights.TAllowTranslateEdit));
            tempxls.SetCellValue(actRow11, 18, getInt(tempUser.userRights.TAllowTranslateRelease));
            tempxls.SetCellValue(actRow11, 19, getInt(tempUser.userRights.TAllowTranslateNew));
            tempxls.SetCellValue(actRow11, 20, getInt(tempUser.userRights.TAllowTranslateDelete));
            actRow11++;

            tempxls.ActiveSheet = 12;
            foreach (int tempRight in tempUser.taskroleRights)
            {
                tempxls.SetCellValue(actRow12, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow12, 2, tempRight);
                actRow12++;
            }

            tempxls.ActiveSheet = 14;
            foreach (string tempRight in tempUser.usergroupRights)
            {
                tempxls.SetCellValue(actRow14, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow14, 2, tempRight);
                actRow14++;
            }

            tempxls.ActiveSheet = 15;
            foreach (int tempRight in tempUser.translateLanguagesRights)
            {
                tempxls.SetCellValue(actRow15, 1, tempUser.UserID);
                tempxls.SetCellValue(actRow15, 2, tempRight);
                actRow15++;
            }
        }
        dataReader.close();
        # endregion LoginUserAndUnitOrgID

        #region Usergroups
        tempxls.ActiveSheet = 13;
        actRow = 2;
        // Schleife über alle DB-Einträge
        SessionObj.ProcessMessage = "UserGroups ausgeben ... ";
        dataReader = new SqlDB("Select groupid FROM loginusergroups ORDER BY groupid", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            string tempValue = dataReader.getString(0);
            // Wert eintragen
            tempxls.SetCellValue(actRow, 1, tempValue);
            actRow++;
        }
        dataReader.close();
        #endregion

        # region UserLanguage
        // Ausgabe der DB-Einträge - eine Zeile pro Wert, die Sprachen in den Spalten
        tempxls.ActiveSheet = 16;
        tempxls.SheetName = "UserLanguage";
        actRow = 1;
        int actValue = -1;
        // 
        tempxls.SetCellValue(actRow, 1, "value");
        tempxls.SetCellValue(actRow, 2, "code");
        // Anzahl der gefunden Sprachen
        int languageCount = 0;
        // Schleife über alle DB-Einträge
        SessionObj.ProcessMessage = "UserLanguages ausgeben ... ";
        dataReader = new SqlDB("Select value, code, text, language FROM loginuserlanguage ORDER BY value, language", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            //Spalte für die Sprache ermitteln und ggf. anlegen
            string test = dataReader.getString(3);
            int colNumber = getColNumber2(dataReader.getString(3), languageCount, tempxls);
            if (colNumber == -1)
            {
                // Sprache ist noch nicht vorhanden -> neue Spalte anlegen
                languageCount++;
                tempxls.SetCellValue(1, languageCount + 2, dataReader.getString(3));
                colNumber = languageCount + 2;
            }
            // Zeile für den Wert ermitteln; die DB-Werte werden sortiert gelesen: änderung des DB-Wertes -> neue Zeile
            int tempValue = dataReader.getInt32(0);
            if (tempValue != actValue)
            {
                actRow++;
                actValue = tempValue;
            }
            // Wert und Text der Sprache eintragen
            tempxls.SetCellValue(actRow, 1, actValue);
            tempxls.SetCellValue(actRow, 2, dataReader.getString(1));
            tempxls.SetCellValue(actRow, colNumber, dataReader.getString(2));
        }
        dataReader.close();
        # endregion LoginLanguage

        # region UserProfiles
        SessionObj.ProcessMessage = "UserProfiles ausgeben ... ";
        
        tempxls.ActiveSheet = 17;
        tempxls.SheetName = "OrgManagerProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "O.EditLevel");
        tempxls.SetCellValue(actRow, 4, "O.AllowDelegate");
        tempxls.SetCellValue(actRow, 5, "O.AllowLock");
        tempxls.SetCellValue(actRow, 6, "O.AllowEmployeeRead");
        tempxls.SetCellValue(actRow, 7, "O.AllowEmployeeEdit");
        tempxls.SetCellValue(actRow, 8, "O.AllowEmployeeImport");
        tempxls.SetCellValue(actRow, 9, "O.AllowEmployeeExport");
        tempxls.SetCellValue(actRow, 10, "O.AllowUnitAdd");
        tempxls.SetCellValue(actRow, 11, "O.AllowUnitMove");
        tempxls.SetCellValue(actRow, 12, "O.AllowUnitDelete");
        tempxls.SetCellValue(actRow, 13, "O.AllowUnitProperty");
        tempxls.SetCellValue(actRow, 14, "O.AllowReportRecipient");
        tempxls.SetCellValue(actRow, 15, "O.AllowStructureImport");
        tempxls.SetCellValue(actRow, 16, "O.AllowStructureExport");
        tempxls.SetCellValue(actRow, 17, "O.AllowBouncesRead");
        tempxls.SetCellValue(actRow, 18, "O.AllowBouncesEdit");
        tempxls.SetCellValue(actRow, 19, "O.AllowBouncesDelete");
        tempxls.SetCellValue(actRow, 20, "O.AllowBouncesExport");
        tempxls.SetCellValue(actRow, 21, "O.AllowBouncesImport");
        tempxls.SetCellValue(actRow, 22, "O.AllowReInvitation");
        tempxls.SetCellValue(actRow, 23, "O.AllowPostNomination");
        tempxls.SetCellValue(actRow, 24, "O.AllowPostNominationImport");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, OAllowBouncesRead, OAllowBouncesEdit, OAllowBouncesDelete, OAllowBouncesExport, OAllowBouncesImport, OAllowReInvitation, OAllowPostNomination, OAllowPostNominationImport FROM orgmanager_profiles ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
            tempxls.SetCellValue(actRow, 17, dataReader.getInt32(16));
            tempxls.SetCellValue(actRow, 18, dataReader.getInt32(17));
            tempxls.SetCellValue(actRow, 19, dataReader.getInt32(18));
            tempxls.SetCellValue(actRow, 20, dataReader.getInt32(19));
            tempxls.SetCellValue(actRow, 21, dataReader.getInt32(20));
            tempxls.SetCellValue(actRow, 22, dataReader.getInt32(21));
            tempxls.SetCellValue(actRow, 23, dataReader.getInt32(22));
            tempxls.SetCellValue(actRow, 24, dataReader.getInt32(23));
        }
        dataReader.close();

        tempxls.ActiveSheet = 18;
        tempxls.SheetName = "OrgManager1Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "O.EditLevel");
        tempxls.SetCellValue(actRow, 4, "O.AllowDelegate");
        tempxls.SetCellValue(actRow, 5, "O.AllowLock");
        tempxls.SetCellValue(actRow, 6, "O.AllowEmployeeRead");
        tempxls.SetCellValue(actRow, 7, "O.AllowEmployeeEdit");
        tempxls.SetCellValue(actRow, 8, "O.AllowEmployeeImport");
        tempxls.SetCellValue(actRow, 9, "O.AllowEmployeeExport");
        tempxls.SetCellValue(actRow, 10, "O.AllowUnitAdd");
        tempxls.SetCellValue(actRow, 11, "O.AllowUnitMove");
        tempxls.SetCellValue(actRow, 12, "O.AllowUnitDelete");
        tempxls.SetCellValue(actRow, 13, "O.AllowUnitProperty");
        tempxls.SetCellValue(actRow, 14, "O.AllowReportRecipient");
        tempxls.SetCellValue(actRow, 15, "O.AllowStructureImport");
        tempxls.SetCellValue(actRow, 16, "O.AllowStructureExport");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport FROM orgmanager_profiles1 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
        }
        dataReader.close();

        tempxls.ActiveSheet = 19;
        tempxls.SheetName = "OrgManager2Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "O.EditLevel");
        tempxls.SetCellValue(actRow, 4, "O.AllowDelegate");
        tempxls.SetCellValue(actRow, 5, "O.AllowLock");
        tempxls.SetCellValue(actRow, 6, "O.AllowEmployeeRead");
        tempxls.SetCellValue(actRow, 7, "O.AllowEmployeeEdit");
        tempxls.SetCellValue(actRow, 8, "O.AllowEmployeeImport");
        tempxls.SetCellValue(actRow, 9, "O.AllowEmployeeExport");
        tempxls.SetCellValue(actRow, 10, "O.AllowUnitAdd");
        tempxls.SetCellValue(actRow, 11, "O.AllowUnitMove");
        tempxls.SetCellValue(actRow, 12, "O.AllowUnitDelete");
        tempxls.SetCellValue(actRow, 13, "O.AllowUnitProperty");
        tempxls.SetCellValue(actRow, 14, "O.AllowReportRecipient");
        tempxls.SetCellValue(actRow, 15, "O.AllowStructureImport");
        tempxls.SetCellValue(actRow, 16, "O.AllowStructureExport");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport FROM orgmanager_profiles2 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
        }
        dataReader.close();

        tempxls.ActiveSheet = 20;
        tempxls.SheetName = "OrgManager3Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "O.EditLevel");
        tempxls.SetCellValue(actRow, 4, "O.AllowDelegate");
        tempxls.SetCellValue(actRow, 5, "O.AllowLock");
        tempxls.SetCellValue(actRow, 6, "O.AllowEmployeeRead");
        tempxls.SetCellValue(actRow, 7, "O.AllowEmployeeEdit");
        tempxls.SetCellValue(actRow, 8, "O.AllowEmployeeImport");
        tempxls.SetCellValue(actRow, 9, "O.AllowEmployeeExport");
        tempxls.SetCellValue(actRow, 10, "O.AllowUnitAdd");
        tempxls.SetCellValue(actRow, 11, "O.AllowUnitMove");
        tempxls.SetCellValue(actRow, 12, "O.AllowUnitDelete");
        tempxls.SetCellValue(actRow, 13, "O.AllowUnitProperty");
        tempxls.SetCellValue(actRow, 14, "O.AllowReportRecipient");
        tempxls.SetCellValue(actRow, 15, "O.AllowStructureImport");
        tempxls.SetCellValue(actRow, 16, "O.AllowStructureExport");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport FROM orgmanager_profiles3 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
        }
        dataReader.close();

        tempxls.ActiveSheet = 21;
        tempxls.SheetName = "OrgManagerBackProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel FROM orgmanager_profilesback ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 22;
        tempxls.SheetName = "OrgManager1BackProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel FROM orgmanager_profiles1back ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 23;
        tempxls.SheetName = "OrgManager2BackProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel FROM orgmanager_profiles2back ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 24;
        tempxls.SheetName = "OrgManager3BackProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, OReadLevel FROM orgmanager_profiles3back ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 25;
        tempxls.SheetName = "ResponseProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, RReadLevel FROM response_profiles ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 26;
        tempxls.SheetName = "Response1Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, RReadLevel FROM response_profiles1 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 27;
        tempxls.SheetName = "Response2Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, RReadLevel FROM response_profiles2 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 28;
        tempxls.SheetName = "Response3Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "O.ReadLevel");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, RReadLevel FROM response_profiles3 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
        }
        dataReader.close();

        tempxls.ActiveSheet = 29;
        tempxls.SheetName = "DownloadProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "D.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "D.AllowMail");
        tempxls.SetCellValue(actRow, 4, "D.AllowDownload");
        tempxls.SetCellValue(actRow, 5, "D.AllowLog");
        tempxls.SetCellValue(actRow, 6, "D.AllowType1");
        tempxls.SetCellValue(actRow, 7, "D.AllowType2");
        tempxls.SetCellValue(actRow, 8, "D.AllowType3");
        tempxls.SetCellValue(actRow, 9, "D.AllowType4");
        tempxls.SetCellValue(actRow, 10, "D.AllowType5");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
        }
        dataReader.close();

        tempxls.ActiveSheet = 30;
        tempxls.SheetName = "Download1Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "D.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "D.AllowMail");
        tempxls.SetCellValue(actRow, 4, "D.AllowDownload");
        tempxls.SetCellValue(actRow, 5, "D.AllowLog");
        tempxls.SetCellValue(actRow, 6, "D.AllowType1");
        tempxls.SetCellValue(actRow, 7, "D.AllowType2");
        tempxls.SetCellValue(actRow, 8, "D.AllowType3");
        tempxls.SetCellValue(actRow, 9, "D.AllowType4");
        tempxls.SetCellValue(actRow, 10, "D.AllowType5");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles1 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
        }
        dataReader.close();

        tempxls.ActiveSheet = 31;
        tempxls.SheetName = "Download2Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "D.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "D.AllowMail");
        tempxls.SetCellValue(actRow, 4, "D.AllowDownload");
        tempxls.SetCellValue(actRow, 5, "D.AllowLog");
        tempxls.SetCellValue(actRow, 6, "D.AllowType1");
        tempxls.SetCellValue(actRow, 7, "D.AllowType2");
        tempxls.SetCellValue(actRow, 8, "D.AllowType3");
        tempxls.SetCellValue(actRow, 9, "D.AllowType4");
        tempxls.SetCellValue(actRow, 10, "D.AllowType5");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles2 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
        }
        dataReader.close();

        tempxls.ActiveSheet = 32;
        tempxls.SheetName = "Download3Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "D.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "D.AllowMail");
        tempxls.SetCellValue(actRow, 4, "D.AllowDownload");
        tempxls.SetCellValue(actRow, 5, "D.AllowLog");
        tempxls.SetCellValue(actRow, 6, "D.AllowType1");
        tempxls.SetCellValue(actRow, 7, "D.AllowType2");
        tempxls.SetCellValue(actRow, 8, "D.AllowType3");
        tempxls.SetCellValue(actRow, 9, "D.AllowType4");
        tempxls.SetCellValue(actRow, 10, "D.AllowType5");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles3 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
        }
        dataReader.close();

        tempxls.ActiveSheet = 33;
        tempxls.SheetName = "FollowUpProfiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "F.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "F.EditLevel");
        tempxls.SetCellValue(actRow, 4, "F.SumLevel");
        tempxls.SetCellValue(actRow, 5, "F.AllowCommunications");
        tempxls.SetCellValue(actRow, 6, "F.AllowMeasures");
        tempxls.SetCellValue(actRow, 7, "F.AllowDelete");
        tempxls.SetCellValue(actRow, 8, "F.AllowExcelExport");
        tempxls.SetCellValue(actRow, 9, "F.AllowReminderOnMeasure");
        tempxls.SetCellValue(actRow, 10, "F.AllowReminderOnUnit");
        tempxls.SetCellValue(actRow, 11, "F.AllowUpUnitsMeasures");
        tempxls.SetCellValue(actRow, 12, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(actRow, 13, "F.AllowThemes");
        tempxls.SetCellValue(actRow, 14, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(actRow, 15, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(actRow, 16, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(actRow, 17, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(actRow, 18, "F.AllowOrders");
        tempxls.SetCellValue(actRow, 19, "F.AllowReminderOnOrderUnit");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
            tempxls.SetCellValue(actRow, 17, dataReader.getInt32(16));
            tempxls.SetCellValue(actRow, 18, dataReader.getInt32(17));
            tempxls.SetCellValue(actRow, 19, dataReader.getInt32(18));
        }
        dataReader.close();

        tempxls.ActiveSheet = 34;
        tempxls.SheetName = "FollowUp1Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "F.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "F.EditLevel");
        tempxls.SetCellValue(actRow, 4, "F.SumLevel");
        tempxls.SetCellValue(actRow, 5, "F.AllowCommunications");
        tempxls.SetCellValue(actRow, 6, "F.AllowMeasures");
        tempxls.SetCellValue(actRow, 7, "F.AllowDelete");
        tempxls.SetCellValue(actRow, 8, "F.AllowExcelExport");
        tempxls.SetCellValue(actRow, 9, "F.AllowReminderOnMeasure");
        tempxls.SetCellValue(actRow, 10, "F.AllowReminderOnUnit");
        tempxls.SetCellValue(actRow, 11, "F.AllowUpUnitsMeasures");
        tempxls.SetCellValue(actRow, 12, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(actRow, 13, "F.AllowThemes");
        tempxls.SetCellValue(actRow, 14, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(actRow, 15, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(actRow, 16, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(actRow, 17, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(actRow, 18, "F.AllowOrders");
        tempxls.SetCellValue(actRow, 19, "F.AllowReminderOnOrderUnit");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles1 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
            tempxls.SetCellValue(actRow, 17, dataReader.getInt32(16));
            tempxls.SetCellValue(actRow, 18, dataReader.getInt32(17));
            tempxls.SetCellValue(actRow, 19, dataReader.getInt32(18));
        }
        dataReader.close();

        tempxls.ActiveSheet = 35;
        tempxls.SheetName = "FollowUp2Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "F.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "F.EditLevel");
        tempxls.SetCellValue(actRow, 4, "F.SumLevel");
        tempxls.SetCellValue(actRow, 5, "F.AllowCommunications");
        tempxls.SetCellValue(actRow, 6, "F.AllowMeasures");
        tempxls.SetCellValue(actRow, 7, "F.AllowDelete");
        tempxls.SetCellValue(actRow, 8, "F.AllowExcelExport");
        tempxls.SetCellValue(actRow, 9, "F.AllowReminderOnMeasure");
        tempxls.SetCellValue(actRow, 10, "F.AllowReminderOnUnit");
        tempxls.SetCellValue(actRow, 11, "F.AllowUpUnitsMeasures");
        tempxls.SetCellValue(actRow, 12, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(actRow, 13, "F.AllowThemes");
        tempxls.SetCellValue(actRow, 14, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(actRow, 15, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(actRow, 16, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(actRow, 17, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(actRow, 18, "F.AllowOrders");
        tempxls.SetCellValue(actRow, 19, "F.AllowReminderOnOrderUnit");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles2 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
            tempxls.SetCellValue(actRow, 17, dataReader.getInt32(16));
            tempxls.SetCellValue(actRow, 18, dataReader.getInt32(17));
            tempxls.SetCellValue(actRow, 19, dataReader.getInt32(18));
        }
        dataReader.close();

        tempxls.ActiveSheet = 36;
        tempxls.SheetName = "FollowUp3Profiles";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "F.ReadLevel");
        tempxls.SetCellValue(actRow, 3, "F.EditLevel");
        tempxls.SetCellValue(actRow, 4, "F.SumLevel");
        tempxls.SetCellValue(actRow, 5, "F.AllowCommunications");
        tempxls.SetCellValue(actRow, 6, "F.AllowMeasures");
        tempxls.SetCellValue(actRow, 7, "F.AllowDelete");
        tempxls.SetCellValue(actRow, 8, "F.AllowExcelExport");
        tempxls.SetCellValue(actRow, 9, "F.AllowReminderOnMeasure");
        tempxls.SetCellValue(actRow, 10, "F.AllowReminderOnUnit");
        tempxls.SetCellValue(actRow, 11, "F.AllowUpUnitsMeasures");
        tempxls.SetCellValue(actRow, 12, "F.AllowColleaguesMeasures");
        tempxls.SetCellValue(actRow, 13, "F.AllowThemes");
        tempxls.SetCellValue(actRow, 14, "F.AllowReminderOnTheme");
        tempxls.SetCellValue(actRow, 15, "F.AllowReminderOnThemeUnit");
        tempxls.SetCellValue(actRow, 16, "F.AllowTopUnitsThemes");
        tempxls.SetCellValue(actRow, 17, "F.AllowColleaguesThemes");
        tempxls.SetCellValue(actRow, 18, "F.AllowOrders");
        tempxls.SetCellValue(actRow, 19, "F.AllowReminderOnOrderUnit");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles3 ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
            tempxls.SetCellValue(actRow, 13, dataReader.getInt32(12));
            tempxls.SetCellValue(actRow, 14, dataReader.getInt32(13));
            tempxls.SetCellValue(actRow, 15, dataReader.getInt32(14));
            tempxls.SetCellValue(actRow, 16, dataReader.getInt32(15));
            tempxls.SetCellValue(actRow, 17, dataReader.getInt32(16));
            tempxls.SetCellValue(actRow, 18, dataReader.getInt32(17));
            tempxls.SetCellValue(actRow, 19, dataReader.getInt32(18));
        }
        dataReader.close();

        tempxls.ActiveSheet = 37;
        tempxls.SheetName = "TeamspaceProfilesTasks";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "T.AllowTaskRead");
        tempxls.SetCellValue(actRow, 3, "T.AllowTaskCreateEditDelete");
        tempxls.SetCellValue(actRow, 4, "T.AllowTaskFileUpload");
        tempxls.SetCellValue(actRow, 5, "T.AllowTaskExport");
        tempxls.SetCellValue(actRow, 6, "T.AllowTaskImport");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, AllowTaskRead, AllowTaskCreateEditDelete, AllowTaskFileUpload, AllowTaskExport, AllowTaskImport FROM teamspace_profilestasks ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
        }
        dataReader.close();

        tempxls.ActiveSheet = 38;
        tempxls.SheetName = "TeamspaceProfilesContacts";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "T.AllowContactRead");
        tempxls.SetCellValue(actRow, 3, "T.AllowContactCreateEditDelete");
        tempxls.SetCellValue(actRow, 4, "T.AllowContactExport");
        tempxls.SetCellValue(actRow, 5, "T.AllowContactImport");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, AllowContactRead, AllowContactCreateEditDelete, AllowContactExport, AllowContactImport FROM teamspace_profilescontacts ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
        }
        dataReader.close();

        tempxls.ActiveSheet = 39;
        tempxls.SheetName = "TeamspaceProfilesContainer";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "T.AllowUpload");
        tempxls.SetCellValue(actRow, 3, "T.AllowDownload");
        tempxls.SetCellValue(actRow, 4, "T.AllowDeleteOwnFiles");
        tempxls.SetCellValue(actRow, 5, "T.AllowDeleteAllFiles");
        tempxls.SetCellValue(actRow, 6, "T.AllowAccessOwnFilesWithoutPassword");
        tempxls.SetCellValue(actRow, 7, "T.AllowAccessAllFilesWithoutPassword");
        tempxls.SetCellValue(actRow, 8, "T.AllowCreateFolder");
        tempxls.SetCellValue(actRow, 9, "T.AllowDeleteOwnFolder");
        tempxls.SetCellValue(actRow, 10, "T.AllowDeleteAllFolder");
        tempxls.SetCellValue(actRow, 11, "T.AllowResetPassword");
        tempxls.SetCellValue(actRow, 12, "T.AllowTakeOwnership");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, allowUpload, allowDownload, allowDeleteOwnFiles, allowDeleteAllFiles, allowAccessOwnFilesWithoutPassword, allowAccessAllFilesWithoutPassword, allowCreateFolder, allowDeleteOwnFolder, allowDeleteAllFolder, allowResetPassword, allowTakeOwnership FROM teamspace_profilescontainer ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
            tempxls.SetCellValue(actRow, 5, dataReader.getInt32(4));
            tempxls.SetCellValue(actRow, 6, dataReader.getInt32(5));
            tempxls.SetCellValue(actRow, 7, dataReader.getInt32(6));
            tempxls.SetCellValue(actRow, 8, dataReader.getInt32(7));
            tempxls.SetCellValue(actRow, 9, dataReader.getInt32(8));
            tempxls.SetCellValue(actRow, 10, dataReader.getInt32(9));
            tempxls.SetCellValue(actRow, 11, dataReader.getInt32(10));
            tempxls.SetCellValue(actRow, 12, dataReader.getInt32(11));
        }
        dataReader.close();

        tempxls.ActiveSheet = 40;
        tempxls.SheetName = "TeamspaceProfilesTranslate";
        actRow = 1;
        tempxls.SetCellValue(actRow, 1, "ProfileID");
        tempxls.SetCellValue(actRow, 2, "T.AllowTranslateRead");
        tempxls.SetCellValue(actRow, 3, "T.AllowTranslateEdit");
        tempxls.SetCellValue(actRow, 4, "T.AllowTranslateRelease");
        // Schleife über alle DB-Einträge
        dataReader = new SqlDB("Select profileID, AllowTranslateRead, AllowTranslateEdit, AllowTranslateRelease FROM teamspace_profilestranslate ORDER BY profileID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            actRow++;
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getInt32(2));
            tempxls.SetCellValue(actRow, 4, dataReader.getInt32(3));
        }
        dataReader.close();
        #endregion UserProfiles

        // Startseite aktivieren
        tempxls.ActiveSheet = 1;

        // Excel generieren
        SessionObj.ProcessMessage = "Excel streamen ... ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDUserExport, 6, 0, 0, 0, SessionObj.Project.ProjectID);
        DateTime actDate = DateTime.Now.Date;
        MemoryStream outStream = new MemoryStream();
        tempxls.Save(outStream, TFileFormats.Xlsx);

        SessionObj.ExportConfig.outStream = outStream;

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDUserExport, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    # endregion User-Export

    //-------------------- Threads für Struktur-Export -------------------------------
    # region StructureExport
    int xslFormatDate;
    protected int getColNumber(string aColName, int languageCount, ExcelFile aXlsFile)
    {
        int colNumber = -1;
        for (int Index = 1; Index <= languageCount; Index++)
        {
            string cellValue = aXlsFile.GetCellValue(1, Index + 1).ToString();
            if (aXlsFile.GetCellValue(1, Index + 1).ToString() == aColName)
                colNumber = Index + 1;
        }
        return colNumber;
    }
    public void doExport()
    {
        // Mitarbeiteranzahl neu berechnen
        SessionObj.ProcessMessage = "Mitarbeiteranzahlen neu berechnen ... ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 2, 0, 0, 0, SessionObj.Project.ProjectID);
        TAdminStructure.recalc(SessionObj.Project.ProjectID);
        TAdminStructure1.recalc(SessionObj.Project.ProjectID);
        TAdminStructure2.recalc(SessionObj.Project.ProjectID);

        // Excel Datei mit 34 Registern erzeugen (Config, Struktur, Mitarbeiter, Codebuch)
        ExcelFile tempxls = new XlsFile(true);
        tempxls.NewFile(35, TExcelFileFormat.v2010);
        TFlxFormat xlsFormat = tempxls.GetDefaultFormat;
        xlsFormat.Format = "dd/mm/YYYY";
        xslFormatDate = tempxls.AddFormat(xlsFormat);

        #region Config
        tempxls.ActiveSheet = 1;
        tempxls.SheetName = "Config";
        int actRow = 1;
        // Header erstellen
        tempxls.SetCellValue(actRow, 1, "ID");
        tempxls.SetCellValue(actRow, 2, "value");
        tempxls.SetCellValue(actRow, 3, "text");
        actRow++;
        // Schleife über alle DB-Einträge
        SqlDB dataReader = new SqlDB("Select ID, value, text FROM orgmanager_config ORDER BY ID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            tempxls.SetCellValue(actRow, 1, dataReader.getString(0));
            tempxls.SetCellValue(actRow, 2, dataReader.getInt32(1));
            tempxls.SetCellValue(actRow, 3, dataReader.getString(2));
            actRow++;
        }
        dataReader.close();
        #endregion Config

        int value;
        TAdminStructure tempStructure = new TAdminStructure(-1, null, 0, SessionObj.Project.ProjectID);

        # region Struktur
        // Struktur ausgeben
        SessionObj.ProcessMessage = "Struktur ausgeben ... ";
        ArrayList orgPath = new ArrayList();
        // Header erstellen
        tempxls.ActiveSheet = 2;
        tempxls.SheetName = "Structure";
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Status");
        for (int index = 0; index < 20; index++)
        {
            tempxls.SetCellValue(1, 57 + index, "delevel" + index.ToString());
        }

        tempxls.SetCellValue(2, 1, tempStructure.OrgID);
        tempxls.SetCellValue(2, 2, tempStructure.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructure.OrgID);
        tempxls.SetCellValue(2, 23, tempStructure.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructure.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructure.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructure.status);

        int actCol = 27;
        if (SessionObj.Project.OrgManagerConfig.ShowOrgEmployeeInputCount)
        {
            tempxls.SetCellValue(1, actCol, "Anzahl eingetragener Mitarbeiter");
            tempxls.SetCellValue(2, actCol, tempStructure.numberEmployeesInput);
            actCol++;
        }
        if (SessionObj.Project.OrgManagerConfig.ShowOrgEmployeeInputCountTotal)
        {
            tempxls.SetCellValue(1, actCol, "Anzahl eingetragener Mitarbeiter aggregiert");
            tempxls.SetCellValue(2, actCol, tempStructure.numberEmployeesInputSum);
            actCol++;
        }
        if (SessionObj.Project.OrgManagerConfig.ShowEmployeeSupervisor)
        {
            tempxls.SetCellValue(1, actCol, "Anzahl Führungskräfte");
            tempxls.SetCellValue(2, actCol, tempStructure.countFK);
            actCol++;
        }
        // Schleife über alle Zeilen
        for (int divIndex = 0; divIndex < SessionObj.Project.structureFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structureFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structureFieldsList.colCount; colIndex++)
                {
                    if (SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                    {
                        TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                        if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                        {
                            switch (tempField.fieldType)
                            {
                                case "UnitName":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "UnitNameShort":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "EmployeeCount":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure.numberEmployees);
                                    // aggregierte Werte
                                    actCol++;
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title + " (+)");
                                    tempxls.SetCellValue(2, actCol, tempStructure.numberEmployeesSum);
                                    actCol++;
                                    break;
                                case "EmployeeCountTotal":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure.numberEmployeesTotal);
                                    actCol++;
                                    break;
                                case "Text":
                                case "Textbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "Date":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellFormat(2, actCol, xslFormatDate);
                                    tempxls.SetCellValue(2, actCol, tempStructure.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "BackComparison":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure.backComparison);
                                    actCol++;
                                    break;
                                case "LanguageUser":
                                case "Radio":
                                case "Dropdown":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                    actCol++;
                                    break;
                                case "Checkbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    string values = "";
                                    ArrayList tempValues = tempStructure.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                    foreach (string tempValue in tempValues)
                                    {
                                        if (values == "")
                                            values = tempValue;
                                        else
                                            values = values + " " + tempValue;
                                    }
                                    tempxls.SetCellValue(2, actCol, values);
                                    actCol++;
                                    break;
                                default:
                                    actCol++;
                                    break;
                            }
                        }
                    }
                }
            }
        tempxls.SetCellValue(1, actCol, "Sperre");
        value = 0;
        if (tempStructure.locked)
            value = 1;
        tempxls.SetCellValue(2, actCol, value);

        actCol++;
        tempxls.SetCellValue(1, actCol, "Workshop");
        if (tempStructure.workshop)
            tempxls.SetCellValue(2, actCol, 1);
        else
            tempxls.SetCellValue(2, actCol, 0);

        actCol++;
        tempxls.SetCellValue(1, actCol, "Kommunikation");
        if (tempStructure.communicationfinished)
            tempxls.SetCellValue(2, actCol, 1);
        else
            tempxls.SetCellValue(2, actCol, 0);
        actCol++;
        tempxls.SetCellValue(1, actCol, "Kommunikationsdatum");
        tempxls.SetCellValue(2, actCol, tempStructure.communicationfinsheddate.ToShortDateString());
        actCol++;
        tempxls.SetCellValue(1, actCol, "Teilnehmeranzahl");
        tempxls.SetCellValue(2, actCol, tempStructure.numberOfParticipants.ToString());

        orgPath.Add(tempStructure.OrgID.ToString());
        actRow = 2;
        outLevel(tempStructure, 1, SessionObj.Project.ProjectID, ref actRow, tempxls, ref orgPath);
        #endregion Struktur

        # region Mitarbeiter
        // Mitarbeiter zusammenstellen
        SessionObj.ProcessMessage = "Mitarbeiter ausgeben ... ";
        tempxls.ActiveSheet = 3;
        tempxls.SheetName = "Employees";

        #region Header erstellen
        actCol = 1;
        tempxls.SetCellValue(1, actCol, "EmployeeID");
        actCol++;
        tempxls.SetCellValue(1, actCol, "OrgID");
        actCol++;
        tempxls.SetCellValue(1, actCol, "AccessCode");
        actCol++;
        // Schleife über alle Zeilen
        for (int divIndex = 0; divIndex < SessionObj.Project.employeeFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.employeeFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.employeeFieldsList.colCount; colIndex++)
                {
                    if (SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex] != null)
                    {
                        TEmployeeFieldsList.TEmployeeFieldsListEntry tempField = (TEmployeeFieldsList.TEmployeeFieldsListEntry)SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex];
                        if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                        {
                            tempxls.SetCellValue(1, actCol, tempField.fieldID);
                            actCol++;
                        }
                    }
                }
            }
        tempxls.SetCellValue(1, actCol, "isBounce");
        actCol++;
        for (int index = 0; index < 20; index++)
        {
            tempxls.SetCellValue(1, actCol + index + 1, "delevel" + index.ToString());
        }
        #endregion Header erstellen

        // Mitarbeiter ausgeben
        actRow = 2;
        orgPath = new ArrayList();
        int actOrgID = 0;
        TEmployee tempEmployee;
        bool isValid = true;
        dataReader = new SqlDB("SELECT employeeID, orgID FROM orgmanager_employee ORDER BY orgID", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            int tempOrgID = dataReader.getInt32(1);
            SqlDB dataReader1 = new SqlDB("select orgID from structure where orgID='" + tempOrgID + "'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                if ((actOrgID == 0) || (dataReader.getInt32(1) != actOrgID) || !isValid)
                {
                    actOrgID = dataReader.getInt32(1);
                    isValid = TAdminStructure.getTopOrgIDList(actOrgID, ref orgPath, SessionObj.Project.ProjectID);
                }
                if (isValid)
                {
                    TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 4, actRow, 0, 0, SessionObj.Project.ProjectID);
                    try
                    {
                        // Mitarbeiter erzeugen
                        int test = dataReader.getInt32(0);
                        tempEmployee = new TEmployee(test, SessionObj.Project.ProjectID);
                        actCol = 1;
                        tempxls.SetCellValue(actRow, actCol, tempEmployee.EmployeeID);
                        actCol++;
                        tempxls.SetCellValue(actRow, actCol, tempEmployee.orgID);
                        actCol++;
                        tempxls.SetCellValue(actRow, actCol, tempEmployee.accesscode);
                        actCol++;
                        // Schleife über alle Zeilen
                        for (int divIndex = 0; divIndex < SessionObj.Project.employeeFieldsList.divisionCount; divIndex++)
                        {
                            for (int rowIndex = 0; rowIndex < SessionObj.Project.employeeFieldsList.rowCount; rowIndex++)
                            {
                                // Schleife über alle Elemente der Zeile
                                for (int colIndex = 0; colIndex < SessionObj.Project.employeeFieldsList.colCount; colIndex++)
                                {
                                    if (SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex] != null)
                                    {
                                        TEmployeeFieldsList.TEmployeeFieldsListEntry tempField = (TEmployeeFieldsList.TEmployeeFieldsListEntry)SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex];
                                        switch (tempField.fieldType)
                                        {
                                            case "NameEmployee":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.Name);
                                                actCol++;
                                                break;
                                            case "FirstnameEmployee":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.Firstname);
                                                actCol++;
                                                break;
                                            case "UnitManager":
                                                if (tempEmployee.isSupervisor)
                                                    tempxls.SetCellValue(actRow, actCol, 1);
                                                else
                                                    tempxls.SetCellValue(actRow, actCol, 0);
                                                actCol++;
                                                break;
                                            case "OrgIDOrigin":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDOrigin);
                                                actCol++;
                                                break;
                                            case "OrgIDOrigin1":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDAdd1Origin);
                                                actCol++;
                                                break;
                                            case "OrgIDOrigin2":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDAdd2Origin);
                                                actCol++;
                                                break;
                                            case "OrgIDOrigin3":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDAdd3Origin);
                                                actCol++;
                                                break;
                                            case "Matrix1":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDAdd1);
                                                actCol++;
                                                break;
                                            case "Matrix2":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDAdd2);
                                                actCol++;
                                                break;
                                            case "Matrix3":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.orgIDAdd3);
                                                actCol++;
                                                break;
                                            case "EmailEmployee":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.Email);
                                                actCol++;
                                                break;
                                            case "MobileEmployee":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.Mobile);
                                                actCol++;
                                                break;
                                            case "Text":
                                            case "Textbox":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.getValue(tempField.fieldID));
                                                actCol++;
                                                break;
                                            case "Date":
                                                tempxls.SetCellFormat(actRow, actCol, xslFormatDate);
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.getValue(tempField.fieldID));
                                                actCol++;
                                                break;
                                            case "Radio":
                                            case "Dropdown":
                                                tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempEmployee.getValue(tempField.fieldID)));
                                                actCol++;
                                                break;
                                            case "Checkbox":
                                                string values = "";
                                                ArrayList tempValues = tempEmployee.getValueList(((TEmployeeFieldsList.TEmployeeFieldsListEntry)(SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                                foreach (string tempValue in tempValues)
                                                {
                                                    if (values == "")
                                                        values = tempValue;
                                                    else
                                                        values = values + " " + tempValue;
                                                }
                                                tempxls.SetCellValue(actRow, actCol, values);
                                                actCol++;
                                                break;
                                            case "SurveyType":
                                                tempxls.SetCellValue(actRow, actCol, tempEmployee.surveyType);
                                                actCol++;
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        if (tempEmployee.isBounce)
                            tempxls.SetCellValue(actRow, actCol, 1);
                        else
                            tempxls.SetCellValue(actRow, actCol, 0);
                        actCol++;
                        for (int index = 0; index < 20; index++)
                        {
                            if (index < orgPath.Count)
                                tempxls.SetCellValue(actRow, actCol + 1 + index, orgPath[index].ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                    actRow++;
                }
            }
            dataReader1.close();
        }
        dataReader.close();
        # endregion Mitarbeiter

        # region Struktur1
        // Struktur1 zusammenstellen
        SessionObj.ProcessMessage = "Struktur1 ausgeben ... ";
        tempxls.ActiveSheet = 4;
        tempxls.SheetName = "Structure1";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Status");
        for (int index = 0; index < 20; index++)
        {
            tempxls.SetCellValue(1, 57 + index, "delevel" + index.ToString());
        }

        // Struktur1 ausgeben
        TAdminStructure1 tempStructure1 = new TAdminStructure1(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructure1.OrgID);
        tempxls.SetCellValue(2, 2, tempStructure1.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructure1.OrgID);
        tempxls.SetCellValue(2, 23, tempStructure1.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructure1.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructure1.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructure1.status);

        actCol = 27;
        // Schleife über alle Zeilen
        for (int divIndex = 0; divIndex < SessionObj.Project.structure1FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure1FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure1FieldsList.colCount; colIndex++)
                {
                    if (SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                    {
                        TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                        if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                        {
                            switch (tempField.fieldType)
                            {
                                case "UnitName":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "UnitNameShort":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "EmployeeCount":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure1.numberEmployees);
                                    // aggregierte Werte
                                    actCol++;
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title + " (+)");
                                    tempxls.SetCellValue(2, actCol, tempStructure1.numberEmployeesSum);
                                    actCol++;
                                    break;
                                case "EmployeeCountTotal":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure1.numberEmployeesTotal);
                                    actCol++;
                                    break;
                                case "Text":
                                case "Textbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure1.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "Date":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellFormat(2, actCol, xslFormatDate);
                                    tempxls.SetCellValue(2, actCol, tempStructure1.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "BackComparison":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure1.backComparison);
                                    actCol++;
                                    break;
                                case "LanguageUser":
                                case "Radio":
                                case "Dropdown":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, Convert.ToInt32(tempStructure1.getValue(tempField.fieldID)));
                                    actCol++;
                                    break;
                                case "Checkbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    string values = "";
                                    ArrayList tempValues = tempStructure1.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                    foreach (string tempValue in tempValues)
                                    {
                                        if (values == "")
                                            values = tempValue;
                                        else
                                            values = values + " " + tempValue;
                                    }
                                    tempxls.SetCellValue(2, actCol, values);
                                    actCol++;
                                    break;
                                default:
                                    actCol++;
                                    break;
                            }
                        }
                    }
                }
            }
        tempxls.SetCellValue(1, actCol, "Sperre");
        value = 0;
        if (tempStructure.locked)
            value = 1;
        tempxls.SetCellValue(2, actCol, value);

        orgPath.Add(tempStructure1.OrgID.ToString());
        actRow = 2;
        outLevel1(tempStructure1, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion Struktur1

        # region Struktur2
        // Struktur2 zusammenstellen
        SessionObj.ProcessMessage = "Struktur2 ausgeben ... ";
        tempxls.ActiveSheet = 5;
        tempxls.SheetName = "Structure2";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Status");
        for (int index = 0; index < 20; index++)
        {
            tempxls.SetCellValue(1, 57 + index, "delevel" + index.ToString());
        }

        // Struktur2 ausgeben
        TAdminStructure2 tempStructure2 = new TAdminStructure2(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructure2.OrgID);
        tempxls.SetCellValue(2, 2, tempStructure2.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructure2.OrgID);
        tempxls.SetCellValue(2, 23, tempStructure2.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructure2.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructure2.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructure2.status);
        actCol = 27;
        // Schleife über alle Zeilen
        for (int divIndex = 0; divIndex < SessionObj.Project.structure2FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure2FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure2FieldsList.colCount; colIndex++)
                {
                    if (SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                    {
                        TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                        if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                        {
                            switch (tempField.fieldType)
                            {
                                case "UnitName":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "UnitNameShort":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "EmployeeCount":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure2.numberEmployees);
                                    // aggregierte Werte
                                    actCol++;
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title + " (+)");
                                    tempxls.SetCellValue(2, actCol, tempStructure2.numberEmployeesSum);
                                    actCol++;
                                    break;
                                case "EmployeeCountTotal":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure2.numberEmployeesTotal);
                                    actCol++;
                                    break;
                                case "Text":
                                case "Textbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure2.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "Date":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellFormat(2, actCol, xslFormatDate);
                                    tempxls.SetCellValue(2, actCol, tempStructure2.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "BackComparison":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure2.backComparison);
                                    actCol++;
                                    break;
                                case "LanguageUser":
                                case "Radio":
                                case "Dropdown":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, Convert.ToInt32(tempStructure2.getValue(tempField.fieldID)));
                                    actCol++;
                                    break;
                                case "Checkbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    string values = "";
                                    ArrayList tempValues = tempStructure2.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                    foreach (string tempValue in tempValues)
                                    {
                                        if (values == "")
                                            values = tempValue;
                                        else
                                            values = values + " " + tempValue;
                                    }
                                    tempxls.SetCellValue(2, actCol, values);
                                    actCol++;
                                    break;
                                default:
                                    actCol++;
                                    break;
                            }
                        }
                    }
                }
            }
        tempxls.SetCellValue(1, actCol, "Sperre");
        value = 0;
        if (tempStructure.locked)
            value = 1;
        tempxls.SetCellValue(2, actCol, value);

        orgPath.Add(tempStructure2.OrgID.ToString());
        actRow = 2;

        outLevel2(tempStructure2, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion Struktur2

        # region Struktur3
        // Struktur3 zusammenstellen
        SessionObj.ProcessMessage = "Struktur3 ausgeben ... ";
        tempxls.ActiveSheet = 6;
        tempxls.SheetName = "Structure3";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Status");
        for (int index = 0; index < 20; index++)
        {
            tempxls.SetCellValue(1, 57 + index, "delevel" + index.ToString());
        }

        // Struktur3 ausgeben
        TAdminStructure3 tempStructure3 = new TAdminStructure3(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructure3.OrgID);
        tempxls.SetCellValue(2, 2, tempStructure3.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructure3.OrgID);
        tempxls.SetCellValue(2, 23, tempStructure3.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructure3.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructure3.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructure3.status);
        actCol = 27;
        // Schleife über alle Zeilen
        for (int divIndex = 0; divIndex < SessionObj.Project.structure3FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure3FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure3FieldsList.colCount; colIndex++)
                {
                    if (SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                    {
                        TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                        if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                        {
                            switch (tempField.fieldType)
                            {
                                case "UnitName":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "UnitNameShort":
                                    // wird in vorherigen Spalten bereits ausgegeben
                                    break;
                                case "EmployeeCount":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure3.numberEmployees);
                                    // aggregierte Werte
                                    actCol++;
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title + " (+)");
                                    tempxls.SetCellValue(2, actCol, tempStructure3.numberEmployeesSum);
                                    actCol++;
                                    break;
                                case "EmployeeCountTotal":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure3.numberEmployeesTotal);
                                    actCol++;
                                    break;
                                case "Text":
                                case "Textbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure3.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "Date":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellFormat(2, actCol, xslFormatDate);
                                    tempxls.SetCellValue(2, actCol, tempStructure3.getValue(tempField.fieldID));
                                    actCol++;
                                    break;
                                case "BackComparison":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, tempStructure3.backComparison);
                                    actCol++;
                                    break;
                                case "LanguageUser":
                                case "Radio":
                                case "Dropdown":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    tempxls.SetCellValue(2, actCol, Convert.ToInt32(tempStructure3.getValue(tempField.fieldID)));
                                    actCol++;
                                    break;
                                case "Checkbox":
                                    tempxls.SetCellValue(1, actCol, Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempField.fieldID, SessionObj.Language, SessionObj.Project.ProjectID).title);
                                    string values = "";
                                    ArrayList tempValues = tempStructure3.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                    foreach (string tempValue in tempValues)
                                    {
                                        if (values == "")
                                            values = tempValue;
                                        else
                                            values = values + " " + tempValue;
                                    }
                                    tempxls.SetCellValue(2, actCol, values);
                                    actCol++;
                                    break;
                                default:
                                    actCol++;
                                    break;
                            }
                        }
                    }
                }
            }
        tempxls.SetCellValue(1, actCol, "Sperre");
        value = 0;
        if (tempStructure.locked)
            value = 1;
        tempxls.SetCellValue(2, actCol, value);

        orgPath.Add(tempStructure3.OrgID.ToString());
        actRow = 2;

        outLevel3(tempStructure3, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion Struktur3

        # region StructureBack
        // Struktur2 zusammenstellen
        SessionObj.ProcessMessage = "StrukturBack ausgeben ... ";
        tempxls.ActiveSheet = 7;
        tempxls.SheetName = "StructureBack";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Filter");
        // StrukturBack ausgeben
        TAdminStructureBack tempStructureBack = new TAdminStructureBack(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructureBack.OrgID);
        tempxls.SetCellValue(2, 2, tempStructureBack.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructureBack.OrgID);
        tempxls.SetCellValue(2, 23, tempStructureBack.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructureBack.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructureBack.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructureBack.filter);
        actRow = 2;
        outLevelBack(tempStructureBack, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion StrukturBack

        # region StructureBack1
        // Struktur2 zusammenstellen
        SessionObj.ProcessMessage = "StrukturBack1 ausgeben ... ";
        tempxls.ActiveSheet = 8;
        tempxls.SheetName = "StructureBack1";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Filter");
        // StrukturBack ausgeben
        TAdminStructureBack1 tempStructureBack1 = new TAdminStructureBack1(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructureBack1.OrgID);
        tempxls.SetCellValue(2, 2, tempStructureBack1.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructureBack1.OrgID);
        tempxls.SetCellValue(2, 23, tempStructureBack1.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructureBack1.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructureBack1.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructureBack1.filter);
        actRow = 2;
        outLevelBack1(tempStructureBack1, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion StrukturBack1

        # region StructureBack2
        // Struktur2 zusammenstellen
        SessionObj.ProcessMessage = "StrukturBack2 ausgeben ... ";
        tempxls.ActiveSheet = 9;
        tempxls.SheetName = "StructureBack2";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Filter");
        // StrukturBack ausgeben
        TAdminStructureBack2 tempStructureBack2 = new TAdminStructureBack2(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructureBack2.OrgID);
        tempxls.SetCellValue(2, 2, tempStructureBack2.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructureBack2.OrgID);
        tempxls.SetCellValue(2, 23, tempStructureBack2.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructureBack2.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructureBack2.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructureBack2.filter);
        actRow = 2;
        outLevelBack2(tempStructureBack2, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion StrukturBack2

        # region StructureBack3
        // Struktur3 zusammenstellen
        SessionObj.ProcessMessage = "StrukturBack3 ausgeben ... ";
        tempxls.ActiveSheet = 10;
        tempxls.SheetName = "StructureBack3";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "Struktur");
        tempxls.SetCellValue(1, 22, "OrgID");
        tempxls.SetCellValue(1, 23, "topOrgID");
        tempxls.SetCellValue(1, 24, "Anzeigename");
        tempxls.SetCellValue(1, 25, "Anzeigename kurz");
        tempxls.SetCellValue(1, 26, "Filter");
        // StrukturBack ausgeben
        TAdminStructureBack3 tempStructureBack3 = new TAdminStructureBack3(-1, null, 0, SessionObj.Project.ProjectID);
        tempxls.SetCellValue(2, 1, tempStructureBack3.OrgID);
        tempxls.SetCellValue(2, 2, tempStructureBack3.orgDisplayName);
        tempxls.SetCellValue(2, 22, tempStructureBack3.OrgID);
        tempxls.SetCellValue(2, 23, tempStructureBack3.TopOrgID);
        tempxls.SetCellValue(2, 24, tempStructureBack3.orgDisplayName);
        tempxls.SetCellValue(2, 25, tempStructureBack3.orgDisplayNameShort);
        tempxls.SetCellValue(2, 26, tempStructureBack3.filter);
        actRow = 2;
        outLevelBack3(tempStructureBack3, 1, SessionObj.Project.ProjectID, ref actRow, tempxls);
        # endregion StrukturBack3

        #region Inputconfig
        SessionObj.ProcessMessage = "Konfiguration ausgeben ... ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 5, 0, 0, 0, SessionObj.Project.ProjectID);

        #region Strukturdefinition
        tempxls.ActiveSheet = 11;
        tempxls.SheetName = "StructureFields";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "division");
        tempxls.SetCellValue(1, 3, "postionRow");
        tempxls.SetCellValue(1, 4, "positionCol");
        tempxls.SetCellValue(1, 5, "fieldType");
        tempxls.SetCellValue(1, 6, "recipient");
        tempxls.SetCellValue(1, 7, "mandatory");
        tempxls.SetCellValue(1, 8, "maxchar");
        tempxls.SetCellValue(1, 9, "width");
        tempxls.SetCellValue(1, 10, "rows");
        tempxls.SetCellValue(1, 11, "minValue");
        tempxls.SetCellValue(1, 12, "maxValue");
        tempxls.SetCellValue(1, 13, "regex");
        // Schleife über alle Zeilen
        actRow = 2;
        for (int divIndex = 0; divIndex < SessionObj.Project.structureFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structureFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structureFieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        tempxls.SetCellValue(actRow, 1, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldID);
                        tempxls.SetCellValue(actRow, 2, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].division);
                        tempxls.SetCellValue(actRow, 3, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionRow);
                        tempxls.SetCellValue(actRow, 4, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionCol);
                        tempxls.SetCellValue(actRow, 5, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType);
                        if (SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].isRecipient)
                            tempxls.SetCellValue(actRow, 6, 1);
                        else
                            tempxls.SetCellValue(actRow, 6, 0);
                        if (SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].madantory)
                            tempxls.SetCellValue(actRow, 7, 1);
                        else
                            tempxls.SetCellValue(actRow, 7, 0);
                        tempxls.SetCellValue(actRow, 8, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxchar);
                        tempxls.SetCellValue(actRow, 9, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].width);
                        tempxls.SetCellValue(actRow, 10, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].rows);
                        if (SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType == "Date")
                        {
                            try
                            {
                                tempxls.SetCellFormat(actRow, 11, xslFormatDate);
                                tempxls.SetCellValue(actRow, 11, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect);
                                tempxls.SetCellFormat(actRow, 12, xslFormatDate);
                                tempxls.SetCellValue(actRow, 12, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect);
                            }
                            catch
                            {
                            }
                        }
                        else
                        {
                            try
                            {
                                tempxls.SetCellValue(actRow, 11, Convert.ToInt32(SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect));
                                tempxls.SetCellValue(actRow, 12, Convert.ToInt32(SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect));
                            }
                            catch
                            {
                            }
                        }
                        tempxls.SetCellValue(actRow, 13, SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].regex);
                        actRow++;
                    }
                    //tempxls.SetCellFormat(2, actCol, xslFormatDate);
                }
            }
        #endregion Strukturdefinition
        #region Strukturdefinitionsfilter
        tempxls.ActiveSheet = 12;
        tempxls.SheetName = "StructureFieldsFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "orgID");
        tempxls.SetCellValue(1, 3, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        TStructureFieldsFilter tempFilterList = new TStructureFieldsFilter("orgmanager_structurefieldsfilter", SessionObj.Project.ProjectID);
        foreach (TStructureFieldsFilter.TEntry tempEntry in tempFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 3, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturdefinitionsfilter
        #region Strukturdefinitionskategorien
        // Header
        tempxls.ActiveSheet = 13;
        tempxls.SheetName = "StructureFieldsCategories";
        tempxls.SetCellValue(1, 1, "fieldID");
        tempxls.SetCellValue(1, 2, "value");
        actRow = 2;
        actCol = 3;
        TLanguages Language = new TLanguages(SessionObj.Project.ProjectID);
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structureFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structureFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structureFieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if ((tempField.fieldType == "Radio") || (tempField.fieldType == "Dropdown") || (tempField.fieldType == "Checkbox"))
                        {
                            // Schleife über alle Kategorien
                            dataReader = new SqlDB("SELECT DISTINCT(value) from orgmanager_structurefieldscategories WHERE fieldID='" + tempField.fieldID + "' ORDER BY value", SessionObj.Project.ProjectID);
                            while (dataReader.read())
                            {
                                tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                                tempxls.SetCellValue(actRow, 2, dataReader.getInt32(0));

                                // Schleife über alle Sprachen
                                actCol = 3;
                                foreach (TLanguages.TEntry tempLanguage in Language.Language)
                                {
                                    SqlDB dataReader1 = new SqlDB("SELECT text from orgmanager_structurefieldscategories WHERE fieldID='" + tempField.fieldID + "' AND value='" + dataReader.getInt32(0) + "' AND language='" + tempLanguage.Language + "'", SessionObj.Project.ProjectID);
                                    if (dataReader1.read())
                                        tempxls.SetCellValue(actRow, actCol, dataReader1.getString(0));
                                    else
                                        tempxls.SetCellValue(actRow, actCol, "");
                                    dataReader1.close();

                                    actCol++;
                                }
                                actRow++;
                            }
                            dataReader.close();
                        }
                    }
                }
            }
        #endregion  Strukturdefinitionskategorien
        #region Strukturdefinitionsfilter
        tempxls.ActiveSheet = 14;
        tempxls.SheetName = "StructureFieldsCategoriesFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "value");
        tempxls.SetCellValue(1, 3, "orgID");
        tempxls.SetCellValue(1, 4, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        TStructureCategoriesFilter tempCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structurefieldscategoriesfilter", SessionObj.Project.ProjectID);
        foreach (TStructureCategoriesFilter.TEntry tempEntry in tempCategoriesFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.value);
            tempxls.SetCellValue(actRow, 3, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 4, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturdefinitionsfilter
        #region Dictionary
        // Header
        tempxls.ActiveSheet = 15;
        tempxls.SheetName = "StructureFieldsDictionary";
        tempxls.SetCellValue(1, 1, "fieldID");
        actRow = 2;
        actCol = 2;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
            tempxls.SetCellValue(1, actCol, "tooltipp");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRegEx");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMandatory");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRange");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMaxLength");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorInvalidData");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHintHeader");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint1");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint2");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint3");
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structureFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structureFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structureFieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                    if (tempField != null)
                    {
                        if (tempField.fieldType != "Empty")
                        {
                            tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                            // Schleife über alle Sprachen
                            actCol = 2;
                            foreach (TLanguages.TEntry tempLanguage in Language.Language)
                            {
                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempField.fieldID, tempLanguage.Language, SessionObj.Project.ProjectID);
                                tempxls.SetCellValue(actRow, actCol, tempLabel.title);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.tooltipp);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRegEx);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMandatory);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRange);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMaxLength);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorInvalidValue);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHintHeader);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint1);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint2);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint3);
                                actCol++;
                            }
                            actRow++;
                        }
                    }
                }
            }
        #endregion Dictionary

        #region Struktur1definition
        tempxls.ActiveSheet = 16;
        tempxls.SheetName = "Structure1Fields";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "division");
        tempxls.SetCellValue(1, 3, "postionRow");
        tempxls.SetCellValue(1, 4, "positionCol");
        tempxls.SetCellValue(1, 5, "fieldType");
        tempxls.SetCellValue(1, 6, "recipient");
        tempxls.SetCellValue(1, 7, "mandatory");
        tempxls.SetCellValue(1, 8, "maxchar");
        tempxls.SetCellValue(1, 9, "width");
        tempxls.SetCellValue(1, 10, "rows");
        tempxls.SetCellValue(1, 11, "minValue");
        tempxls.SetCellValue(1, 12, "maxValue");
        tempxls.SetCellValue(1, 13, "regex");
        // Schleife über alle Zeilen
        actRow = 2;
        for (int divIndex = 0; divIndex < SessionObj.Project.structure1FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure1FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure1FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                    if (tempField != null)
                    {
                        tempxls.SetCellValue(actRow, 1, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldID);
                        tempxls.SetCellValue(actRow, 2, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].division);
                        tempxls.SetCellValue(actRow, 3, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionRow);
                        tempxls.SetCellValue(actRow, 4, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionCol);
                        tempxls.SetCellValue(actRow, 5, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType);
                        if (SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].isRecipient)
                            tempxls.SetCellValue(actRow, 6, 1);
                        else
                            tempxls.SetCellValue(actRow, 6, 0);
                        if (SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].madantory)
                            tempxls.SetCellValue(actRow, 7, 1);
                        else
                            tempxls.SetCellValue(actRow, 7, 0);
                        tempxls.SetCellValue(actRow, 8, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxchar);
                        tempxls.SetCellValue(actRow, 9, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].width);
                        tempxls.SetCellValue(actRow, 10, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].rows);
                        if (SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType == "Date")
                        {
                            try
                            {
                                tempxls.SetCellFormat(actRow, 11, xslFormatDate);
                                tempxls.SetCellValue(actRow, 11, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect);
                                tempxls.SetCellFormat(actRow, 12, xslFormatDate);
                                tempxls.SetCellValue(actRow, 12, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect);
                            }
                            catch
                            {
                            }
                        }
                        else
                        {
                            try
                            {
                                tempxls.SetCellValue(actRow, 11, Convert.ToInt32(SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect));
                                tempxls.SetCellValue(actRow, 12, Convert.ToInt32(SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect));
                            }
                            catch
                            {
                            }
                        }
                        tempxls.SetCellValue(actRow, 13, SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].regex);
                        actRow++;
                    }
                }
            }
        #endregion Struktur1definition
        #region Struktur1definitionsfilter
        tempxls.ActiveSheet = 17;
        tempxls.SheetName = "Structure1FieldsFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "orgID");
        tempxls.SetCellValue(1, 3, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempFilterList = new TStructureFieldsFilter("orgmanager_structure1fieldsfilter", SessionObj.Project.ProjectID);
        foreach (TStructureFieldsFilter.TEntry tempEntry in tempFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 3, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturde1finitionsfilter
        #region Struktur1definitionskategorien
        // Header
        tempxls.ActiveSheet = 18;
        tempxls.SheetName = "Structure1FieldsCategories";
        tempxls.SetCellValue(1, 1, "fieldID");
        tempxls.SetCellValue(1, 2, "value");
        actRow = 2;
        actCol = 3;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structure1FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure1FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure1FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if ((tempField.fieldType == "Radio") || (tempField.fieldType == "Dropdown") || (tempField.fieldType == "Checkbox"))
                        {
                            // Schleife über alle Kategorien
                            dataReader = new SqlDB("SELECT DISTINCT(value) from orgmanager_structure1fieldscategories ORDER BY value", SessionObj.Project.ProjectID);
                            while (dataReader.read())
                            {
                                tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                                tempxls.SetCellValue(actRow, 2, dataReader.getInt32(0));

                                // Schleife über alle Sprachen
                                actCol = 3;
                                foreach (TLanguages.TEntry tempLanguage in Language.Language)
                                {
                                    SqlDB dataReader1 = new SqlDB("SELECT text from orgmanager_structure1fieldscategories WHERE fieldID='" + tempField.fieldID + "' AND value='" + dataReader.getInt32(0) + "' AND language='" + tempLanguage.Language + "'", SessionObj.Project.ProjectID);
                                    if (dataReader1.read())
                                        tempxls.SetCellValue(actRow, actCol, dataReader1.getString(0));
                                    else
                                        tempxls.SetCellValue(actRow, actCol, "");
                                    dataReader1.close();

                                    actCol++;
                                }
                                actRow++;
                            }
                        }
                        dataReader.close();
                    }
                }
            }
        #endregion  Struktur1definitionskategorien
        #region Struktur1definitionsfilter
        tempxls.ActiveSheet = 19;
        tempxls.SheetName = "Structure1FieldsCategoriesFilte";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "value");
        tempxls.SetCellValue(1, 3, "orgID");
        tempxls.SetCellValue(1, 4, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structure1fieldscategoriesfilter", SessionObj.Project.ProjectID);
        foreach (TStructureCategoriesFilter.TEntry tempEntry in tempCategoriesFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.value);
            tempxls.SetCellValue(actRow, 3, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 4, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturde1finitionsfilter
        #region Dictionary
        // Header
        tempxls.ActiveSheet = 20;
        tempxls.SheetName = "Structure1FieldsDictionary";
        tempxls.SetCellValue(1, 1, "fieldID");
        actRow = 2;
        actCol = 2;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
            tempxls.SetCellValue(1, actCol, "tooltipp");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRegEx");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMandatory");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRange");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMaxLength");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorInvalidData");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHintHeader");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint1");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint2");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint3");
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structure1FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure1FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure1FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if (tempField.fieldType != "Empty")
                        {
                            tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                            // Schleife über alle Sprachen
                            actCol = 2;
                            foreach (TLanguages.TEntry tempLanguage in Language.Language)
                            {
                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempField.fieldID, tempLanguage.Language, SessionObj.Project.ProjectID);
                                tempxls.SetCellValue(actRow, actCol, tempLabel.title);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.tooltipp);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRegEx);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMandatory);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRange);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMaxLength);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorInvalidValue);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHintHeader);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint1);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint2);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint3);
                                actCol++;
                            }
                            actRow++;
                        }
                    }
                }
            }
        #endregion Dictionary

        #region Struktur2definition
        tempxls.ActiveSheet = 21;
        tempxls.SheetName = "Structure2Fields";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "division");
        tempxls.SetCellValue(1, 3, "postionRow");
        tempxls.SetCellValue(1, 4, "positionCol");
        tempxls.SetCellValue(1, 5, "fieldType");
        tempxls.SetCellValue(1, 6, "recipient");
        tempxls.SetCellValue(1, 7, "mandatory");
        tempxls.SetCellValue(1, 8, "maxchar");
        tempxls.SetCellValue(1, 9, "width");
        tempxls.SetCellValue(1, 10, "rows");
        tempxls.SetCellValue(1, 11, "minValue");
        tempxls.SetCellValue(1, 12, "maxValue");
        tempxls.SetCellValue(1, 13, "regex");
        // Schleife über alle Zeilen
        actRow = 2;
        for (int divIndex = 0; divIndex < SessionObj.Project.structure2FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure2FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure2FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        tempxls.SetCellValue(actRow, 1, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldID);
                        tempxls.SetCellValue(actRow, 2, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].division);
                        tempxls.SetCellValue(actRow, 3, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionRow);
                        tempxls.SetCellValue(actRow, 4, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionCol);
                        tempxls.SetCellValue(actRow, 5, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType);
                        if (SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].isRecipient)
                            tempxls.SetCellValue(actRow, 6, 1);
                        else
                            tempxls.SetCellValue(actRow, 6, 0);
                        if (SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].madantory)
                            tempxls.SetCellValue(actRow, 7, 1);
                        else
                            tempxls.SetCellValue(actRow, 7, 0);
                        tempxls.SetCellValue(actRow, 8, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxchar);
                        tempxls.SetCellValue(actRow, 9, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].width);
                        tempxls.SetCellValue(actRow, 10, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].rows);
                        if (SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType == "Date")
                        {
                            try
                            {
                                tempxls.SetCellFormat(actRow, 11, xslFormatDate);
                                tempxls.SetCellValue(actRow, 11, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect);
                                tempxls.SetCellFormat(actRow, 12, xslFormatDate);
                                tempxls.SetCellValue(actRow, 12, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect);
                            }
                            catch
                            {
                            }
                        }
                        else
                        {
                            try
                            {
                                tempxls.SetCellValue(actRow, 11, Convert.ToInt32(SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect));
                                tempxls.SetCellValue(actRow, 12, Convert.ToInt32(SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect));
                            }
                            catch
                            {
                            }
                        }
                        tempxls.SetCellValue(actRow, 13, SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].regex);
                        actRow++;
                    }
                }
            }
        #endregion Struktur2definition
        #region Struktur2definitionsfilter
        tempxls.ActiveSheet = 22;
        tempxls.SheetName = "Structure2FieldsFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "orgID");
        tempxls.SetCellValue(1, 3, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempFilterList = new TStructureFieldsFilter("orgmanager_structure2fieldsfilter", SessionObj.Project.ProjectID);
        foreach (TStructureFieldsFilter.TEntry tempEntry in tempFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 3, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturde2finitionsfilter
        #region Struktur2definitionskategorien
        // Header
        tempxls.ActiveSheet = 23;
        tempxls.SheetName = "Structure2FieldsCategories";
        tempxls.SetCellValue(1, 1, "fieldID");
        tempxls.SetCellValue(1, 2, "value");
        actRow = 2;
        actCol = 3;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structure2FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure2FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure2FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if ((tempField.fieldType == "Radio") || (tempField.fieldType == "Dropdown") || (tempField.fieldType == "Checkbox"))
                        {
                            // Schleife über alle Kategorien
                            dataReader = new SqlDB("SELECT DISTINCT(value) from orgmanager_structure2fieldscategories ORDER BY value", SessionObj.Project.ProjectID);
                            while (dataReader.read())
                            {
                                tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                                tempxls.SetCellValue(actRow, 2, dataReader.getInt32(0));

                                // Schleife über alle Sprachen
                                actCol = 3;
                                foreach (TLanguages.TEntry tempLanguage in Language.Language)
                                {
                                    SqlDB dataReader1 = new SqlDB("SELECT text from orgmanager_structure2fieldscategories WHERE fieldID='" + tempField.fieldID + "' AND value='" + dataReader.getInt32(0) + "' AND language='" + tempLanguage.Language + "'", SessionObj.Project.ProjectID);
                                    if (dataReader1.read())
                                        tempxls.SetCellValue(actRow, actCol, dataReader1.getString(0));
                                    else
                                        tempxls.SetCellValue(actRow, actCol, "");
                                    dataReader1.close();

                                    actCol++;
                                }
                                actRow++;
                            }
                        }
                        dataReader.close();
                    }
                }
            }
        #endregion  Struktur2definitionskategorien
        #region Struktur2definitionsfilter
        tempxls.ActiveSheet = 24;
        tempxls.SheetName = "Structure2FieldsCategoriesFilte";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "value");
        tempxls.SetCellValue(1, 3, "orgID");
        tempxls.SetCellValue(1, 4, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structure2fieldscategoriesfilter", SessionObj.Project.ProjectID);
        foreach (TStructureCategoriesFilter.TEntry tempEntry in tempCategoriesFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.value);
            tempxls.SetCellValue(actRow, 3, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 4, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturde2finitionsfilter
        #region Dictionary
        // Header
        tempxls.ActiveSheet = 25;
        tempxls.SheetName = "Structure2FieldsDictionary";
        tempxls.SetCellValue(1, 1, "fieldID");
        actRow = 2;
        actCol = 2;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
            tempxls.SetCellValue(1, actCol, "tooltipp");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRegEx");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMandatory");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRange");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMaxLength");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorInvalidData");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHintHeader");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint1");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint2");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint3");
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structure2FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure2FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure2FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if (tempField.fieldType != "Empty")
                        {
                            tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                            // Schleife über alle Sprachen
                            actCol = 2;
                            foreach (TLanguages.TEntry tempLanguage in Language.Language)
                            {
                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2", tempField.fieldID, tempLanguage.Language, SessionObj.Project.ProjectID);
                                tempxls.SetCellValue(actRow, actCol, tempLabel.title);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.tooltipp);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRegEx);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMandatory);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRange);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMaxLength);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorInvalidValue);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHintHeader);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint1);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint2);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint3);
                                actCol++;
                            }
                            actRow++;
                        }
                    }
                }
            }
        #endregion Dictionary

        #region Struktur3definition
        tempxls.ActiveSheet = 26;
        tempxls.SheetName = "Structure3Fields";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "division");
        tempxls.SetCellValue(1, 3, "postionRow");
        tempxls.SetCellValue(1, 4, "positionCol");
        tempxls.SetCellValue(1, 5, "fieldType");
        tempxls.SetCellValue(1, 6, "recipient");
        tempxls.SetCellValue(1, 7, "mandatory");
        tempxls.SetCellValue(1, 8, "maxchar");
        tempxls.SetCellValue(1, 9, "width");
        tempxls.SetCellValue(1, 10, "rows");
        tempxls.SetCellValue(1, 11, "minValue");
        tempxls.SetCellValue(1, 12, "maxValue");
        tempxls.SetCellValue(1, 13, "regex");
        // Schleife über alle Zeilen
        actRow = 2;
        for (int divIndex = 0; divIndex < SessionObj.Project.structure3FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure3FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure3FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        tempxls.SetCellValue(actRow, 1, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldID);
                        tempxls.SetCellValue(actRow, 2, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].division);
                        tempxls.SetCellValue(actRow, 3, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionRow);
                        tempxls.SetCellValue(actRow, 4, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].positionCol);
                        tempxls.SetCellValue(actRow, 5, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType);
                        if (SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].isRecipient)
                            tempxls.SetCellValue(actRow, 6, 1);
                        else
                            tempxls.SetCellValue(actRow, 6, 0);
                        if (SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].madantory)
                            tempxls.SetCellValue(actRow, 7, 1);
                        else
                            tempxls.SetCellValue(actRow, 7, 0);
                        tempxls.SetCellValue(actRow, 8, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxchar);
                        tempxls.SetCellValue(actRow, 9, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].width);
                        tempxls.SetCellValue(actRow, 10, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].rows);
                        if (SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType == "Date")
                        {
                            try
                            {
                                tempxls.SetCellFormat(actRow, 11, xslFormatDate);
                                tempxls.SetCellValue(actRow, 11, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect);
                                tempxls.SetCellFormat(actRow, 12, xslFormatDate);
                                tempxls.SetCellValue(actRow, 12, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect);
                            }
                            catch
                            {
                            }
                        }
                        else
                        {
                            try
                            {
                                tempxls.SetCellValue(actRow, 11, Convert.ToInt32(SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].minselect));
                                tempxls.SetCellValue(actRow, 12, Convert.ToInt32(SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].maxselect));
                            }
                            catch
                            {
                            }
                        }
                        tempxls.SetCellValue(actRow, 13, SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex].regex);
                        actRow++;
                    }
                }
            }
        #endregion Struktur3definition
        #region Struktur3definitionsfilter
        tempxls.ActiveSheet = 27;
        tempxls.SheetName = "Structure3FieldsFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "orgID");
        tempxls.SetCellValue(1, 3, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempFilterList = new TStructureFieldsFilter("orgmanager_structure3fieldsfilter", SessionObj.Project.ProjectID);
        foreach (TStructureFieldsFilter.TEntry tempEntry in tempFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 3, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturde3finitionsfilter
        #region Struktur3definitionskategorien
        // Header
        tempxls.ActiveSheet = 28;
        tempxls.SheetName = "Structure3FieldsCategories";
        tempxls.SetCellValue(1, 1, "fieldID");
        tempxls.SetCellValue(1, 2, "value");
        actRow = 2;
        actCol = 3;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structure3FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure3FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure3FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if ((tempField.fieldType == "Radio") || (tempField.fieldType == "Dropdown") || (tempField.fieldType == "Checkbox"))
                        {
                            // Schleife über alle Kategorien
                            dataReader = new SqlDB("SELECT DISTINCT(value) from orgmanager_structure3fieldscategories ORDER BY value", SessionObj.Project.ProjectID);
                            while (dataReader.read())
                            {
                                tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                                tempxls.SetCellValue(actRow, 2, dataReader.getInt32(0));

                                // Schleife über alle Sprachen
                                actCol = 3;
                                foreach (TLanguages.TEntry tempLanguage in Language.Language)
                                {
                                    SqlDB dataReader1 = new SqlDB("SELECT text from orgmanager_structure3fieldscategories WHERE fieldID='" + tempField.fieldID + "' AND value='" + dataReader.getInt32(0) + "' AND language='" + tempLanguage.Language + "'", SessionObj.Project.ProjectID);
                                    if (dataReader1.read())
                                        tempxls.SetCellValue(actRow, actCol, dataReader1.getString(0));
                                    else
                                        tempxls.SetCellValue(actRow, actCol, "");
                                    dataReader1.close();

                                    actCol++;
                                }
                                actRow++;
                            }
                        }
                        dataReader.close();
                    }
                }
            }
        #endregion  Struktur3definitionskategorien
        #region Struktur3definitionsfilter
        tempxls.ActiveSheet = 29;
        tempxls.SheetName = "Structure3FieldsCategoriesFilte";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "value");
        tempxls.SetCellValue(1, 3, "orgID");
        tempxls.SetCellValue(1, 4, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structure3fieldscategoriesfilter", SessionObj.Project.ProjectID);
        foreach (TStructureCategoriesFilter.TEntry tempEntry in tempCategoriesFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.value);
            tempxls.SetCellValue(actRow, 3, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 4, tempEntry.righttype);

            actRow++;
        }
        #endregion Strukturde3finitionsfilter
        #region Dictionary
        // Header
        tempxls.ActiveSheet = 30;
        tempxls.SheetName = "Structure3FieldsDictionary";
        tempxls.SetCellValue(1, 1, "fieldID");
        actRow = 2;
        actCol = 2;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
            tempxls.SetCellValue(1, actCol, "tooltipp");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRegEx");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMandatory");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRange");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMaxLength");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorInvalidData");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHintHeader");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint1");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint2");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint3");
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.structure3FieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.structure3FieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.structure3FieldsList.colCount; colIndex++)
                {
                    TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if (tempField.fieldType != "Empty")
                        {
                            tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                            // Schleife über alle Sprachen
                            actCol = 2;
                            foreach (TLanguages.TEntry tempLanguage in Language.Language)
                            {
                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3", tempField.fieldID, tempLanguage.Language, SessionObj.Project.ProjectID);
                                tempxls.SetCellValue(actRow, actCol, tempLabel.title);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.tooltipp);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRegEx);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMandatory);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRange);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMaxLength);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorInvalidValue);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHintHeader);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint1);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint2);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint3);
                                actCol++;
                            }
                            actRow++;
                        }
                    }
                }
            }
        #endregion Dictionary

        #region Employeedefinition
        tempxls.ActiveSheet = 31;
        tempxls.SheetName = "EmployeeFields";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "division");
        tempxls.SetCellValue(1, 3, "postionRow");
        tempxls.SetCellValue(1, 4, "positionCol");
        tempxls.SetCellValue(1, 5, "fieldType");
        tempxls.SetCellValue(1, 6, "mandatory");
        tempxls.SetCellValue(1, 7, "maxchar");
        tempxls.SetCellValue(1, 8, "width");
        tempxls.SetCellValue(1, 9, "rows");
        tempxls.SetCellValue(1, 10, "minValue");
        tempxls.SetCellValue(1, 11, "maxValue");
        tempxls.SetCellValue(1, 12, "regex");
        // Schleife über alle Zeilen
        actRow = 2;
        for (int divIndex = 0; divIndex < SessionObj.Project.employeeFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.employeeFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.employeeFieldsList.colCount; colIndex++)
                {
                    TEmployeeFieldsList.TEmployeeFieldsListEntry tempField = (TEmployeeFieldsList.TEmployeeFieldsListEntry)SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        tempxls.SetCellValue(actRow, 1, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].fieldID);
                        tempxls.SetCellValue(actRow, 2, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].division);
                        tempxls.SetCellValue(actRow, 3, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].positionRow);
                        tempxls.SetCellValue(actRow, 4, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].positionCol);
                        tempxls.SetCellValue(actRow, 5, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].fieldType);
                        if (SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].madantory)
                            tempxls.SetCellValue(actRow, 6, 1);
                        else
                            tempxls.SetCellValue(actRow, 6, 0);
                        tempxls.SetCellValue(actRow, 7, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].maxchar);
                        tempxls.SetCellValue(actRow, 8, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].width);
                        tempxls.SetCellValue(actRow, 9, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].rows);
                        if (SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].fieldType == "Date")
                        {
                            try
                            {
                                tempxls.SetCellFormat(actRow, 10, xslFormatDate);
                                tempxls.SetCellValue(actRow, 10, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].minselect);
                                tempxls.SetCellFormat(actRow, 11, xslFormatDate);
                                tempxls.SetCellValue(actRow, 11, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].maxselect);
                            }
                            catch
                            {
                            }
                        }
                        else
                        {
                            try
                            {
                                tempxls.SetCellValue(actRow, 10, Convert.ToInt32(SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].minselect));
                                tempxls.SetCellValue(actRow, 11, Convert.ToInt32(SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].maxselect));
                            }
                            catch
                            {
                            }
                        }
                        tempxls.SetCellValue(actRow, 12, SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex].regex);
                        actRow++;
                    }
                }
            }
        #endregion Employeedefinition
        #region Employeedefinitionsfilter
        tempxls.ActiveSheet = 32;
        tempxls.SheetName = "EmployeeFieldsFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "orgID");
        tempxls.SetCellValue(1, 3, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempFilterList = new TStructureFieldsFilter("orgmanager_employeefieldsfilter", SessionObj.Project.ProjectID);
        foreach (TStructureFieldsFilter.TEntry tempEntry in tempFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 3, tempEntry.righttype);

            actRow++;
        }
        #endregion Employeefinitionsfilter
        #region Employeedefinitionskategorien
        // Header
        tempxls.ActiveSheet = 33;
        tempxls.SheetName = "EmployeeFieldsCategories";
        tempxls.SetCellValue(1, 1, "fieldID");
        tempxls.SetCellValue(1, 2, "value");
        actRow = 2;
        actCol = 3;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.employeeFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.employeeFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.employeeFieldsList.colCount; colIndex++)
                {
                    TEmployeeFieldsList.TEmployeeFieldsListEntry tempField = (TEmployeeFieldsList.TEmployeeFieldsListEntry)SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if ((tempField.fieldType == "Radio") || (tempField.fieldType == "Dropdown") || (tempField.fieldType == "Checkbox"))
                        {
                            // Schleife über alle Kategorien
                            dataReader = new SqlDB("SELECT DISTINCT(value) from orgmanager_employeefieldscategories WHERE fieldID='" + tempField.fieldID + "' ORDER BY value", SessionObj.Project.ProjectID);
                            while (dataReader.read())
                            {
                                tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                                tempxls.SetCellValue(actRow, 2, dataReader.getInt32(0));

                                // Schleife über alle Sprachen
                                actCol = 3;
                                foreach (TLanguages.TEntry tempLanguage in Language.Language)
                                {

                                    SqlDB dataReader1 = new SqlDB("SELECT text from orgmanager_employeefieldscategories WHERE fieldID='" + tempField.fieldID + "' AND value='" + dataReader.getInt32(0) + "' AND language='" + tempLanguage.Language + "'", SessionObj.Project.ProjectID);
                                    if (dataReader1.read())
                                        tempxls.SetCellValue(actRow, actCol, dataReader1.getString(0));
                                    else
                                        tempxls.SetCellValue(actRow, actCol, "");
                                    dataReader1.close();

                                    actCol++;
                                }
                                actRow++;
                            }
                            dataReader.close();
                        }
                    }
                }
            }
        #endregion  Employeedefinitionskategorien
        #region Employeedefinitionsfilter
        tempxls.ActiveSheet = 34;
        tempxls.SheetName = "EmployeeFieldsCategoriesFilter";
        // Header erstellen
        tempxls.SetCellValue(1, 1, "FieldID");
        tempxls.SetCellValue(1, 2, "value");
        tempxls.SetCellValue(1, 3, "orgID");
        tempxls.SetCellValue(1, 4, "rightType");
        // Schleife über alle Filter
        actRow = 2;
        tempCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_employeefieldscategoriesfilter", SessionObj.Project.ProjectID);
        foreach (TStructureCategoriesFilter.TEntry tempEntry in tempCategoriesFilterList.filterList)
        {
            tempxls.SetCellValue(actRow, 1, tempEntry.fieldID);
            tempxls.SetCellValue(actRow, 2, tempEntry.value);
            tempxls.SetCellValue(actRow, 3, tempEntry.orgID);
            tempxls.SetCellValue(actRow, 4, tempEntry.righttype);

            actRow++;
        }
        #endregion Employeedefinitionsfilter
        #region Dictionary
        // Header
        tempxls.ActiveSheet = 35;
        tempxls.SheetName = "EmployeeFieldsDictionary";
        tempxls.SetCellValue(1, 1, "fieldID");
        actRow = 2;
        actCol = 2;
        foreach (TLanguages.TEntry tempLanguage in Language.Language)
        {
            tempxls.SetCellValue(1, actCol, tempLanguage.Language);
            actCol++;
            tempxls.SetCellValue(1, actCol, "tooltipp");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRegEx");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMandatory");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorRange");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorMaxLength");
            actCol++;
            tempxls.SetCellValue(1, actCol, "errorInvalidData");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHintHeader");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint1");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint2");
            actCol++;
            tempxls.SetCellValue(1, actCol, "exportHint3");
            actCol++;
        }
        // Texte
        for (int divIndex = 0; divIndex < SessionObj.Project.employeeFieldsList.divisionCount; divIndex++)
            for (int rowIndex = 0; rowIndex < SessionObj.Project.employeeFieldsList.rowCount; rowIndex++)
            {
                // Schleife über alle Elemente der Zeile
                for (int colIndex = 0; colIndex < SessionObj.Project.employeeFieldsList.colCount; colIndex++)
                {
                    TEmployeeFieldsList.TEmployeeFieldsListEntry tempField = (TEmployeeFieldsList.TEmployeeFieldsListEntry)SessionObj.Project.employeeFieldsList.employeeFieldMatrix[divIndex, rowIndex, colIndex];

                    if (tempField != null)
                    {
                        if (tempField.fieldType != "Empty")
                        {
                            tempxls.SetCellValue(actRow, 1, tempField.fieldID);
                            // Schleife über alle Sprachen
                            actCol = 2;
                            foreach (TLanguages.TEntry tempLanguage in Language.Language)
                            {
                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_employeefields", tempField.fieldID, tempLanguage.Language, SessionObj.Project.ProjectID);
                                tempxls.SetCellValue(actRow, actCol, tempLabel.title);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.tooltipp);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRegEx);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMandatory);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorRange);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorMaxLength);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.errorInvalidValue);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHintHeader);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint1);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint2);
                                actCol++;
                                tempxls.SetCellValue(actRow, actCol, tempLabel.exportHint3);
                                actCol++;
                            }
                            actRow++;
                        }
                    }
                }
            }
        #endregion Dictionary

        #endregion Inputconfig

        // Startseite aktivieren
        tempxls.ActiveSheet = 1;

        // Excel generieren
        SessionObj.ProcessMessage = "Excel streamen ... ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 6, 0, 0, 0, SessionObj.Project.ProjectID);
        DateTime actDate = DateTime.Now.Date;
        MemoryStream outStream = new MemoryStream();
        tempxls.Save(outStream, TFileFormats.Xlsx);

        SessionObj.ExportConfig.outStream = outStream;

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    protected void fillSheet(int sheetNumber, string sheetName, string tabelName, ExcelFile aXlsFile, string aProjectID)
    {
        // Ausgabe der DB-Einträge - eine Zeile pro Wert, die Sprachen in den Spalten
        aXlsFile.ActiveSheet = sheetNumber;
        aXlsFile.SheetName = sheetName;
        int actRow = 1;
        string actID = "";
        int actValue = 0;
        aXlsFile.SetCellValue(actRow, 1, "fieldID");
        aXlsFile.SetCellValue(actRow, 2, "value");
        // Anzahl der gefunden Sprachen
        int languageCount = 0;
        // Schleife über alle DB-Einträge
        SqlDB dataReader = new SqlDB("SELECT field, value, text, language FROM " + tabelName + " ORDER BY fieldID, value, language", aProjectID);
        while (dataReader.read())
        {
            //Spalte für die Sprache ermitteln und ggf. anlegen
            int colNumber = getColNumber(dataReader.getString(2), languageCount, aXlsFile);
            if (colNumber == -1)
            {
                // Sprache ist noch nicht vorhanden -> neue Spalte anlegen
                languageCount++;
                aXlsFile.SetCellValue(1, languageCount + 2, dataReader.getString(2));
                colNumber = languageCount + 2;
            }
            // Zeile für den Wert ermitteln; die DB-Werte werden sortiert gelesen: änderung des DB-Wertes -> neue Zeile
            string tempID = dataReader.getString(0);
            int tempvalue = dataReader.getInt32(1);
            if (tempID != actID)
            {
                actRow++;
                actID = tempID;
            }
            else
                if (tempvalue != actValue)
            {
                actRow++;
                actValue = tempvalue;
            }
            // Wert und Text der Sprache eintragen
            aXlsFile.SetCellValue(actRow, 2, actValue);
            aXlsFile.SetCellValue(actRow, colNumber, dataReader.getString(2));
        }
        dataReader.close();
    }
    private void outLevel(TAdminStructure aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls, ref ArrayList aOrgPath)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structure where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            int value;
            actRow++;
            TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 3, actRow, 0, 0, SessionObj.Project.ProjectID);
            TAdminStructure tempStructure = new TAdminStructure(dataReader.getInt32(0), null, level, aProjectID);

            tempxls.SetCellValue(actRow, level + 1, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructure.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructure.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructure.status);

            int actCol = 27;
            if (SessionObj.Project.OrgManagerConfig.ShowOrgEmployeeInputCount)
            {
                tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesInput);
                actCol++;
            }
            if (SessionObj.Project.OrgManagerConfig.ShowOrgEmployeeInputCountTotal)
            {
                tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesInputSum);
                actCol++;
            }
            if (SessionObj.Project.OrgManagerConfig.ShowEmployeeSupervisor)
            {
                tempxls.SetCellValue(actRow, actCol, tempStructure.countFK);
                actCol++;
            }
            // Schleife über alle Zeilen
            for (int divIndex = 0; divIndex < SessionObj.Project.structureFieldsList.divisionCount; divIndex++)
                for (int rowIndex = 0; rowIndex < SessionObj.Project.structureFieldsList.rowCount; rowIndex++)
                {
                    // Schleife über alle Elemente der Zeile
                    for (int colIndex = 0; colIndex < SessionObj.Project.structureFieldsList.colCount; colIndex++)
                    {
                        if (SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                        {
                            TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                            if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                            {
                                switch (tempField.fieldType)
                                {
                                    case "UnitName":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "UnitNameShort":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "EmployeeCount":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployees);
                                        // aggregierte Werte
                                        actCol++;
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesSum);
                                        actCol++;
                                        break;
                                    case "EmployeeCountTotal":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesTotal);
                                        actCol++;
                                        break;
                                    case "Text":
                                        if ((tempField.minselect != "") && (tempField.maxselect != ""))
                                        {
                                            string tempValue = tempStructure.getValue(tempField.fieldID);
                                            if (tempValue == "")
                                                tempxls.SetCellValue(actRow, actCol, "");
                                            else
                                                tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        }
                                        else
                                            tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Textbox":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Date":
                                        tempxls.SetCellFormat(actRow, actCol, xslFormatDate);
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "BackComparison":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.backComparison);
                                        actCol++;
                                        break;
                                    case "LanguageUser":
                                    case "Radio":
                                    case "Dropdown":
                                        tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        actCol++;
                                        break;
                                    case "Checkbox":
                                        string values = "";
                                        ArrayList tempValues = tempStructure.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structureFieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                        foreach (string tempValue in tempValues)
                                        {
                                            if (values == "")
                                                values = tempValue;
                                            else
                                                values = values + " " + tempValue;
                                        }
                                        tempxls.SetCellValue(actRow, actCol, values);
                                        actCol++;
                                        break;
                                    default:
                                        actCol++;
                                        break;
                                }
                            }
                        }
                    }
                }
            value = 0;
            if (tempStructure.locked)
                value = 1;
            tempxls.SetCellValue(actRow, actCol, value);

            actCol++;
            if (tempStructure.workshop)
                tempxls.SetCellValue(actRow, actCol, 1);
            else
                tempxls.SetCellValue(actRow, actCol, 0);

            actCol++;
            if (tempStructure.communicationfinished)
                tempxls.SetCellValue(actRow, actCol, 1);
            else
                tempxls.SetCellValue(actRow, actCol, 0);
            actCol++;
            tempxls.SetCellValue(actRow, actCol, tempStructure.communicationfinsheddate.ToShortDateString());
            actCol++;
            tempxls.SetCellValue(actRow, actCol, tempStructure.numberOfParticipants.ToString());

            aOrgPath.Add(tempStructure.OrgID.ToString());
            for (int index = 0; index < 20; index++)
            {
                if (index < aOrgPath.Count)
                    tempxls.SetCellValue(actRow, 60 + index, aOrgPath[index].ToString());
            }
            outLevel(tempStructure, level + 1, aProjectID, ref actRow, tempxls, ref aOrgPath);
            aOrgPath.RemoveAt(aOrgPath.Count - 1);
        }
        dataReader.close();
    }
    private void outLevel1(TAdminStructure1 aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structure1 where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            int value;
            actRow++;
            TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 3, actRow, 0, 0, SessionObj.Project.ProjectID);
            TAdminStructure1 tempStructure = new TAdminStructure1(dataReader.getInt32(0), null, level, aProjectID);

            tempxls.SetCellValue(actRow, level + 1, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructure.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructure.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructure.status);

            int actCol = 27;
            // Schleife über alle Zeilen
            for (int divIndex = 0; divIndex < SessionObj.Project.structure1FieldsList.divisionCount; divIndex++)
                for (int rowIndex = 0; rowIndex < SessionObj.Project.structure1FieldsList.rowCount; rowIndex++)
                {
                    // Schleife über alle Elemente der Zeile
                    for (int colIndex = 0; colIndex < SessionObj.Project.structure1FieldsList.colCount; colIndex++)
                    {
                        if (SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                        {
                            TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                            if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                            {
                                switch (tempField.fieldType)
                                {
                                    case "UnitName":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "UnitNameShort":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "EmployeeCount":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployees);
                                        // aggregierte Werte
                                        actCol++;
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesSum);
                                        actCol++;
                                        break;
                                    case "EmployeeCountTotal":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesTotal);
                                        actCol++;
                                        break;
                                    case "Text":
                                        if ((tempField.minselect != "") && (tempField.maxselect != ""))
                                        {
                                            string tempValue = tempStructure.getValue(tempField.fieldID);
                                            if (tempValue == "")
                                                tempxls.SetCellValue(actRow, actCol, "");
                                            else
                                                tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        }
                                        else
                                            tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Textbox":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Date":
                                        tempxls.SetCellFormat(actRow, actCol, xslFormatDate);
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "BackComparison":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.backComparison);
                                        actCol++;
                                        break;
                                    case "LanguageUser":
                                    case "Radio":
                                    case "Dropdown":
                                        tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        actCol++;
                                        break;
                                    case "Checkbox":
                                        string values = "";
                                        ArrayList tempValues = tempStructure.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structure1FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                        foreach (string tempValue in tempValues)
                                        {
                                            if (values == "")
                                                values = tempValue;
                                            else
                                                values = values + " " + tempValue;
                                        }
                                        tempxls.SetCellValue(actRow, actCol, values);
                                        actCol++;
                                        break;
                                    default:
                                        actCol++;
                                        break;
                                }
                            }
                        }
                    }
                }
            value = 0;
            if (tempStructure.locked)
                value = 1;
            tempxls.SetCellValue(actRow, actCol, value);

            outLevel1(tempStructure, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    private void outLevel2(TAdminStructure2 aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structure2 where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            int value;
            actRow++;
            TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 3, actRow, 0, 0, SessionObj.Project.ProjectID);
            TAdminStructure2 tempStructure = new TAdminStructure2(dataReader.getInt32(0), null, level, aProjectID);

            tempxls.SetCellValue(actRow, level + 1, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructure.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructure.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructure.status);

            int actCol = 27;
            // Schleife über alle Zeilen
            for (int divIndex = 0; divIndex < SessionObj.Project.structure2FieldsList.divisionCount; divIndex++)
                for (int rowIndex = 0; rowIndex < SessionObj.Project.structure2FieldsList.rowCount; rowIndex++)
                {
                    // Schleife über alle Elemente der Zeile
                    for (int colIndex = 0; colIndex < SessionObj.Project.structure2FieldsList.colCount; colIndex++)
                    {
                        if (SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                        {
                            TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                            if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                            {
                                switch (tempField.fieldType)
                                {
                                    case "UnitName":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "UnitNameShort":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "EmployeeCount":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesInput);
                                        // aggregierte Werte
                                        actCol++;
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesSum);
                                        actCol++;
                                        break;
                                    case "EmployeeCountTotal":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesTotal);
                                        actCol++;
                                        break;
                                    case "Text":
                                        if ((tempField.minselect != "") && (tempField.maxselect != ""))
                                        {
                                            string tempValue = tempStructure.getValue(tempField.fieldID);
                                            if (tempValue == "")
                                                tempxls.SetCellValue(actRow, actCol, "");
                                            else
                                                tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        }
                                        else
                                            tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Textbox":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Date":
                                        tempxls.SetCellFormat(actRow, actCol, xslFormatDate);
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "BackComparison":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.backComparison);
                                        actCol++;
                                        break;
                                    case "LanguageUser":
                                    case "Radio":
                                    case "Dropdown":
                                        tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        actCol++;
                                        break;
                                    case "Checkbox":
                                        string values = "";
                                        ArrayList tempValues = tempStructure.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structure2FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                        foreach (string tempValue in tempValues)
                                        {
                                            if (values == "")
                                                values = tempValue;
                                            else
                                                values = values + " " + tempValue;
                                        }
                                        tempxls.SetCellValue(actRow, actCol, values);
                                        actCol++;
                                        break;
                                    default:
                                        actCol++;
                                        break;
                                }
                            }
                        }
                    }
                }
            value = 0;
            if (tempStructure.locked)
                value = 1;
            tempxls.SetCellValue(actRow, actCol, value);

            outLevel2(tempStructure, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    private void outLevel3(TAdminStructure3 aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structure3 where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            int value;
            actRow++;
            TProcessStatus.updateProcessInfo(SessionObj.ProcessIDStructureExport, 3, actRow, 0, 0, SessionObj.Project.ProjectID);
            TAdminStructure3 tempStructure = new TAdminStructure3(dataReader.getInt32(0), null, level, aProjectID);

            tempxls.SetCellValue(actRow, level + 1, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructure.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructure.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructure.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructure.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructure.status);

            int actCol = 27;
            // Schleife über alle Zeilen
            for (int divIndex = 0; divIndex < SessionObj.Project.structure3FieldsList.divisionCount; divIndex++)
                for (int rowIndex = 0; rowIndex < SessionObj.Project.structure3FieldsList.rowCount; rowIndex++)
                {
                    // Schleife über alle Elemente der Zeile
                    for (int colIndex = 0; colIndex < SessionObj.Project.structure3FieldsList.colCount; colIndex++)
                    {
                        if (SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex] != null)
                        {
                            TStructureFieldsList.TStructureFieldsListEntry tempField = (TStructureFieldsList.TStructureFieldsListEntry)SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex];
                            if ((tempField.fieldType != "Label") && (tempField.fieldType != "Empty") && (tempField.fieldType != "Division"))
                            {
                                switch (tempField.fieldType)
                                {
                                    case "UnitName":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "UnitNameShort":
                                        // wird in vorherigen Spalten bereits ausgegeben
                                        break;
                                    case "EmployeeCount":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesInput);
                                        // aggregierte Werte
                                        actCol++;
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesSum);
                                        actCol++;
                                        break;
                                    case "EmployeeCountTotal":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.numberEmployeesTotal);
                                        actCol++;
                                        break;
                                    case "Text":
                                        if ((tempField.minselect != "") && (tempField.maxselect != ""))
                                        {
                                            string tempValue = tempStructure.getValue(tempField.fieldID);
                                            if (tempValue == "")
                                                tempxls.SetCellValue(actRow, actCol, "");
                                            else
                                                tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        }
                                        else
                                            tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Textbox":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "Date":
                                        tempxls.SetCellFormat(actRow, actCol, xslFormatDate);
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.getValue(tempField.fieldID));
                                        actCol++;
                                        break;
                                    case "BackComparison":
                                        tempxls.SetCellValue(actRow, actCol, tempStructure.backComparison);
                                        actCol++;
                                        break;
                                    case "LanguageUser":
                                    case "Radio":
                                    case "Dropdown":
                                        tempxls.SetCellValue(actRow, actCol, Convert.ToInt32(tempStructure.getValue(tempField.fieldID)));
                                        actCol++;
                                        break;
                                    case "Checkbox":
                                        string values = "";
                                        ArrayList tempValues = tempStructure.getValueList(((TStructureFieldsList.TStructureFieldsListEntry)(SessionObj.Project.structure3FieldsList.structureFieldMatrix[divIndex, rowIndex, colIndex])).fieldID);
                                        foreach (string tempValue in tempValues)
                                        {
                                            if (values == "")
                                                values = tempValue;
                                            else
                                                values = values + " " + tempValue;
                                        }
                                        tempxls.SetCellValue(actRow, actCol, values);
                                        actCol++;
                                        break;
                                    default:
                                        actCol++;
                                        break;
                                }
                            }
                        }
                    }
                }
            value = 0;
            if (tempStructure.locked)
                value = 1;
            tempxls.SetCellValue(actRow, actCol, value);

            outLevel3(tempStructure, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    private void outLevelBack(TAdminStructureBack aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structureBack where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            actRow++;
            TAdminStructureBack tempStructureBack = new TAdminStructureBack(dataReader.getInt32(0), null, 0, aProjectID);
            tempxls.SetCellValue(actRow, level + 1, tempStructureBack.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructureBack.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructureBack.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructureBack.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructureBack.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructureBack.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructureBack.filter);

            outLevelBack(tempStructureBack, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    private void outLevelBack1(TAdminStructureBack1 aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structureBack1 where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            actRow++;
            TAdminStructureBack1 tempStructureBack1 = new TAdminStructureBack1(dataReader.getInt32(0), null, 0, aProjectID);
            tempxls.SetCellValue(actRow, level + 1, tempStructureBack1.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructureBack1.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructureBack1.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructureBack1.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructureBack1.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructureBack1.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructureBack1.filter);

            outLevelBack1(tempStructureBack1, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    private void outLevelBack2(TAdminStructureBack2 aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structureBack2 where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            actRow++;
            TAdminStructureBack2 tempStructureBack2 = new TAdminStructureBack2(dataReader.getInt32(0), null, 0, aProjectID);
            tempxls.SetCellValue(actRow, level + 1, tempStructureBack2.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructureBack2.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructureBack2.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructureBack2.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructureBack2.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructureBack2.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructureBack2.filter);

            outLevelBack2(tempStructureBack2, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    private void outLevelBack3(TAdminStructureBack3 aStructure, int level, string aProjectID, ref int actRow, ExcelFile tempxls)
    {
        // Schleife über alle Kinder
        SqlDB dataReader = new SqlDB("SELECT orgID from structureBack3 where topOrgID='" + aStructure.OrgID + "'", aProjectID);
        while (dataReader.read())
        {
            actRow++;
            TAdminStructureBack3 tempStructureBack3 = new TAdminStructureBack3(dataReader.getInt32(0), null, 0, aProjectID);
            tempxls.SetCellValue(actRow, level + 1, tempStructureBack3.OrgID);
            tempxls.SetCellValue(actRow, level + 2, tempStructureBack3.orgDisplayName);
            tempxls.SetCellValue(actRow, 22, tempStructureBack3.OrgID);
            tempxls.SetCellValue(actRow, 23, tempStructureBack3.TopOrgID);
            tempxls.SetCellValue(actRow, 24, tempStructureBack3.orgDisplayName);
            tempxls.SetCellValue(actRow, 25, tempStructureBack3.orgDisplayNameShort);
            tempxls.SetCellValue(actRow, 26, tempStructureBack3.filter);

            outLevelBack3(tempStructureBack3, level + 1, aProjectID, ref actRow, tempxls);
        }
        dataReader.close();
    }
    # endregion StructureExport

    //-------------------- Threads für CollectUser für UserEdit -------------------------------
    # region CollectUser
    /// <summary>
    /// Überprüfung einer Zeichenkette
    /// </summary>
    /// <param name="multiUser">Vergleichswert 1</param>
    /// <param name="tempUser">Vergleichswert 2</param>
    protected void checkValue(ref string multiUser, string tempUser)
    {
        if (multiUser != "*")
        {
            if (multiUser != tempUser)
            {
                multiUser = "*";
            }
        }
    }
    /// <summary>
    /// Überprüfung einer Zeichenkette
    /// </summary>
    /// <param name="multiUser">Vergleichswert 1</param>
    /// <param name="tempUser">Vergleichswert 2</param>
    protected void checkValue(ref int multiUser, int tempUser)
    {
        if (multiUser != -2)
        {
            if (multiUser != tempUser)
            {
                multiUser = -2;
            }
        }
    }
    /// <summary>
    /// Überprüfung einer Zeichenkette
    /// </summary>
    /// <param name="multiUser">Vergleichswert 1</param>
    /// <param name="tempUser">Vergleichswert 2</param>
    protected void checkValue(ref int multiUser, bool tempUser)
    {
        if (multiUser != -1)
        {
            if (((multiUser == 0) & (tempUser)) | ((multiUser == 1) & (!tempUser)))
            {
                multiUser = -1;
            }
        }
    }
    /// <summary>
    /// Überprüfung der Tool-Verfügbarkeit
    /// </summary>
    /// <param name="multiuser">Liste der Berechtigungen</param>
    /// <param name="aToolID">ID des zu überprüfenden Tools</param>
    /// <param name="aAvailable">Verfügbarkeit-Status</param>
    /// <param name="aActice">Aktiv-Status</param>
    protected void checkValue(ref ArrayList multiuser, string aToolID, bool aAvailable)
    {
        foreach (TMultiUser.TToolRight tempTool in multiuser)
        {
            if (tempTool.toolID == aToolID)
            {
                if (((tempTool.available == 0) & (aAvailable)) | ((tempTool.available == 1) & (!aAvailable)))
                {
                    tempTool.available = -1;
                }
            }
        }
    }
    public void doCollectUser()
    {
        // Userlsite erstellen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 1, 0, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);

        TMultiUser multiUser = new TMultiUser();
        TUser tempUser = new TUser((string)(SessionObj.selectedIDList[0]), SessionObj.Project.ProjectID);
        multiUser.UserID = tempUser.UserID;
        multiUser.DisplayName = tempUser.DisplayName;
        multiUser.Password = tempUser.Password;
        multiUser.Email = tempUser.Email;
        multiUser.Title = tempUser.Title;
        multiUser.Name = tempUser.Name;
        multiUser.Firstname = tempUser.Firstname;
        multiUser.Street = tempUser.Street;
        multiUser.Zipcode = tempUser.Zipcode;
        multiUser.City = tempUser.City;
        multiUser.State = tempUser.State;
        multiUser.LastLoginErrorCount = tempUser.LastLoginErrorCount.ToString();
        multiUser.IsInternalUser = 0;
        if (tempUser.IsInternalUser)
            multiUser.IsInternalUser = 1;
        multiUser.ForceChange = 0;
        if (tempUser.ForceChange)
            multiUser.ForceChange = 1;
        multiUser.Active = 0;
        if (tempUser.Active)
            multiUser.Active = 1;
        multiUser.Language = tempUser.Language;
        multiUser.Comment = tempUser.Comment;
        multiUser.GroupID = tempUser.GroupID;
        // Einheiten-Berechtigungen
        multiUser.unitRights = tempUser.orgRights;
        multiUser.maxUnitRightID = tempUser.maxOrgRightID;
        multiUser.backRights = tempUser.backRights;
        multiUser.maxBackRightID = tempUser.maxBackRightID;
        multiUser.org1Rights = tempUser.org1Rights;
        multiUser.maxOrg1RightID = tempUser.maxOrg1RightID;
        multiUser.back1Rights = tempUser.back1Rights;
        multiUser.maxBack1RightID = tempUser.maxBack1RightID;
        multiUser.org2Rights = tempUser.org2Rights;
        multiUser.maxOrg2RightID = tempUser.maxOrg2RightID;
        multiUser.back2Rights = tempUser.back2Rights;
        multiUser.maxBack2RightID = tempUser.maxBack2RightID;
        multiUser.org3Rights = tempUser.org3Rights;
        multiUser.maxOrg3RightID = tempUser.maxOrg3RightID;
        multiUser.back3Rights = tempUser.back3Rights;
        multiUser.maxBack3RightID = tempUser.maxBack3RightID;
        // Container-Berechtigungen
        multiUser.containerRights = tempUser.containerRights;
        multiUser.maxContainerRightID = tempUser.maxContainerRightID;
        // Projektplanberechtigungen
        multiUser.userRights = tempUser.userRights;
        multiUser.taskroleRights = tempUser.taskroleRights;
        multiUser.translateLanguagesRights = tempUser.translateLanguagesRights;
        // UserGroup Berechtigungen
        multiUser.usergroupRights = tempUser.usergroupRights;
        // Tool Berechtigungen
        foreach (TTools.TToolEntry tempTool in SessionObj.tools.tools)
        {
            TMultiUser.TToolRight aTool = new TMultiUser.TToolRight();
            aTool.toolID = tempTool.toolID;
            aTool.available = 0;
            if (tempUser.getToolAvailable(tempTool.toolID))
                aTool.available = 1;
            multiUser.toolRights.Add(aTool);
        }

        for (int Index = 1; Index < SessionObj.selectedIDList.Count; Index++)
        {
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, Index, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);

            tempUser = new TUser((string)(SessionObj.selectedIDList[Index]), SessionObj.Project.ProjectID);
            // Daten vergleichen
            checkValue(ref multiUser.UserID, tempUser.UserID);
            checkValue(ref multiUser.DisplayName, tempUser.DisplayName);
            checkValue(ref multiUser.Password, tempUser.Password);
            checkValue(ref multiUser.Email, tempUser.Email);
            checkValue(ref multiUser.Title, tempUser.Title);
            checkValue(ref multiUser.Name, tempUser.Name);
            checkValue(ref multiUser.Firstname, tempUser.Firstname);
            checkValue(ref multiUser.Street, tempUser.Street);
            checkValue(ref multiUser.Zipcode, tempUser.Zipcode);
            checkValue(ref multiUser.City, tempUser.City);
            checkValue(ref multiUser.State, tempUser.State);
            checkValue(ref multiUser.LastLoginErrorCount, tempUser.LastLoginErrorCount.ToString());
            checkValue(ref multiUser.IsInternalUser, tempUser.IsInternalUser);
            checkValue(ref multiUser.ForceChange, tempUser.ForceChange);
            checkValue(ref multiUser.Active, tempUser.Active);
            checkValue(ref multiUser.Language, tempUser.Language);
            checkValue(ref multiUser.Comment, tempUser.Comment);
            checkValue(ref multiUser.GroupID, tempUser.GroupID);

            foreach (TTools.TToolEntry tempTool in SessionObj.tools.tools)
            {
                checkValue(ref multiUser.toolRights, tempTool.toolID, tempUser.getToolAvailable(tempTool.toolID));
            }
        }
        SessionObj.ActMultiUser = multiUser;

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    #endregion CollectUser

    //-------------------- Threads für CollectRights für UserRightsEdit -------------------------------
    # region CollectRights
    public void doCollectRights()
    {
        // Rightsliste erstellen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 1, 0, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);

        TMultiRights multiRights = new TMultiRights(SessionObj.selectedIDList);

        TRights tempRights = new TRights(Convert.ToInt32(SessionObj.selectedIDList[0].ToString()), SessionObj.selectedIDListTable, SessionObj.userGroups, SessionObj.Project.ProjectID);
        multiRights.UserID = tempRights.UserID;
        multiRights.OrgID = tempRights.OrgID.ToString();
        multiRights.OProfile = tempRights.OProfile;
        multiRights.OReadLevel = tempRights.OReadLevel;
        multiRights.OEditLevel = tempRights.OEditLevel;
        if (tempRights.OAllowDelegate)
            multiRights.OAllowDelegate = 1;
        if (tempRights.OAllowLock)
            multiRights.OAllowLock = 1;
        if (tempRights.OAllowEmployeeRead)
            multiRights.OAllowEmployeeRead = 1;
        if (tempRights.OAllowEmployeeEdit)
            multiRights.OAllowEmployeeEdit = 1;
        if (tempRights.OAllowEmployeeImport)
            multiRights.OAllowEmployeeImport = 1;
        if (tempRights.OAllowEmployeeExport)
            multiRights.OAllowEmployeeExport = 1;
        if (tempRights.OAllowUnitAdd)
            multiRights.OAllowUnitAdd = 1;
        if (tempRights.OAllowUnitMove)
            multiRights.OAllowUnitMove = 1;
        if (tempRights.OAllowUnitDelete)
            multiRights.OAllowUnitDelete = 1;
        if (tempRights.OAllowUnitProperty)
            multiRights.OAllowUnitProperty = 1;
        if (tempRights.OAllowReportRecipient)
            multiRights.OAllowReportRecipient = 1;
        if (tempRights.OAllowStructureImport)
            multiRights.OAllowStructureImport = 1;
        if (tempRights.OAllowStructureExport)
            multiRights.OAllowStructureExport = 1;
        if (tempRights.OAllowBouncesRead)
            multiRights.OAllowBouncesRead = 1;
        if (tempRights.OAllowBouncesEdit)
            multiRights.OAllowBouncesEdit = 1;
        if (tempRights.OAllowBouncesDelete)
            multiRights.OAllowBouncesDelete = 1;
        if (tempRights.OAllowBouncesExport)
            multiRights.OAllowBouncesExport = 1;
        if (tempRights.OAllowBouncesImport)
            multiRights.OAllowBouncesImport = 1;
        if (tempRights.OAllowReInvitation)
            multiRights.OAllowReInvitation = 1;
        if (tempRights.OAllowPostNomination)
            multiRights.OAllowPostNomination = 1;
        if (tempRights.OAllowPostNominationImport)
            multiRights.OAllowPostNominationImport = 1;
        multiRights.FProfile = tempRights.FProfile;
        multiRights.FReadLevel = tempRights.FReadLevel;
        multiRights.FEditLevel = tempRights.FEditLevel;
        multiRights.FSumLevel = tempRights.FSumLevel;
        if (tempRights.FAllowCommunications)
            multiRights.FAllowCommunications = 1;
        if (tempRights.FAllowMeasure)
            multiRights.FAllowMeasures = 1;
        if (tempRights.FAllowDelete)
            multiRights.FAllowDelete = 1;
        if (tempRights.FAllowExcelExport)
            multiRights.FAllowExcelExport = 1;
        if (tempRights.FAllowReminderOnMeasure)
            multiRights.FAllowReminderOnMeasure = 1;
        if (tempRights.FAllowReminderOnUnit)
            multiRights.FAllowReminderOnUnit = 1;
        if (tempRights.FAllowUpUnitsMeasures)
            multiRights.FAllowUpUnitsMeasures = 1;
        if (tempRights.FAllowColleaguesMeasures)
            multiRights.FAllowColleaguesMeasures = 1;
        if (tempRights.FAllowThemes)
            multiRights.FAllowThemes = 1;
        if (tempRights.FAllowReminderOnThemes)
            multiRights.FAllowReminderOnThemes = 1;
        if (tempRights.FAllowReminderOnThemeUnit)
            multiRights.FAllowReminderOnThemeUnit = 1;
        if (tempRights.FAllowColleaguesThemes)
            multiRights.FAllowColleaguesThemes = 1;
        if (tempRights.FAllowOrders)
            multiRights.FAllowOrders = 1;
        if (tempRights.FAllowReminderOnOrderUnit)
            multiRights.FAllowReminderOnOrderUnit = 1;
        multiRights.RProfile = tempRights.RProfile;
        multiRights.RReadLevel = tempRights.RReadLevel;
        multiRights.DProfile = tempRights.DProfile;
        multiRights.DReadLevel = tempRights.DReadLevel;
        if (tempRights.DAllowMail)
            multiRights.DAllowMail = 1;
        if (tempRights.DAllowDownload)
            multiRights.DAllowDownload = 1;
        if (tempRights.DAllowLog)
            multiRights.DAllowLog = 1;
        if (tempRights.DAllowType1)
            multiRights.DAllowType1 = 1;
        if (tempRights.DAllowType2)
            multiRights.DAllowType2 = 1;
        if (tempRights.DAllowType3)
            multiRights.DAllowType3 = 1;
        if (tempRights.DAllowType4)
            multiRights.DAllowType4 = 1;
        if (tempRights.DAllowType5)
            multiRights.DAllowType5 = 1;
        multiRights.AEditLevelWorkshop = tempRights.AEditLevelWorkshop;
        multiRights.AReadLevelStructure = tempRights.AReadLevelStructure;
        multiRights.containerID = tempRights.containerID;
        if (tempRights.TAllowUpload)
            multiRights.TAllowUpload = 1;
        if (tempRights.TAllowDownload)
            multiRights.TAllowDownload = 1;
        if (tempRights.TAllowDeleteOwnFiles)
            multiRights.TAllowDeleteOwnFiles = 1;
        if (tempRights.TAllowDeleteAllFiles)
            multiRights.TAllowDeleteAllFiles = 1;
        if (tempRights.TAllowAccessOwnFilesWithoutPassword)
            multiRights.TAllowAccessOwnFilesWithoutPassword = 1;
        if (tempRights.TAllowAccessAllFilesWithoutPassword)
            multiRights.TAllowAccessAllFilesWithoutPassword = 1;
        if (tempRights.TAllowCreateFolder)
            multiRights.TAllowCreateFolder = 1;
        if (tempRights.TAllowDeleteOwnFiles)
            multiRights.TAllowDeleteOwnFolder = 1;
        if (tempRights.TAllowDeleteAllFolder)
            multiRights.TAllowDeleteAllFolder = 1;
        if (tempRights.TAllowResetPassword)
            multiRights.TAllowResetPassword = 1;
        if (tempRights.TAllowTakeOwnership)
            multiRights.TAllowTakeOwnership = 1;
        if (tempRights.TAllowTaskRead)
            multiRights.TAllowTaskRead = 1;
        if (tempRights.TAllowTaskEdit)
            multiRights.TAllowTaskEdit = 1;
        if (tempRights.TAllowTaskFileUpload)
            multiRights.TAllowTaskFileUpload = 1;
        if (tempRights.TAllowTaskExport)
            multiRights.TAllowTaskExport = 1;
        if (tempRights.TAllowTaskImport)
            multiRights.TAllowTaskImport = 1;
        if (tempRights.TAllowContactRead)
            multiRights.TAllowContactRead = 1;
        if (tempRights.TAllowContactEdit)
            multiRights.TAllowContactEdit = 1;
        if (tempRights.TAllowContactExport)
            multiRights.TAllowContactExport = 1;
        if (tempRights.TAllowContactImport)
            multiRights.TAllowContactImport = 1;

        if (tempRights.AAllowAccountTransfer)
            multiRights.AAllowAccountTransfer = 1;
        if (tempRights.AAllowAccountManagement)
            multiRights.AAllowAccountManagement = 1;

        multiRights.groupID = tempRights.groupID;

        if (tempRights.TAllowTranslateRead)
            multiRights.TAllowTranslateRead = 1;
        if (tempRights.TAllowTranslateEdit)
            multiRights.TAllowTranslateEdit = 1;
        if (tempRights.TAllowTranslateRealease)
            multiRights.TAllowTranslateRelease = 1;
        if (tempRights.TAllowTranslateNew)
            multiRights.TAllowTranslateNew = 1;
        if (tempRights.TAllowTranslateDelete)
            multiRights.TAllowTranslateDelete = 1;
        multiRights.roleID = tempRights.roleID;
        multiRights.translateLanaguageID = tempRights.translateLanguageID;

        for (int Index = 1; Index < SessionObj.selectedIDList.Count; Index++)
        {
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, Index, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);

            tempRights = new TRights(Convert.ToInt32(SessionObj.selectedIDList[Index].ToString()), SessionObj.selectedIDListTable, SessionObj.userGroups, SessionObj.Project.ProjectID);
            // Daten vergleichen
            checkValue(ref multiRights.UserID, tempRights.UserID);
            checkValue(ref multiRights.OrgID, tempRights.OrgID.ToString());
            checkValue(ref multiRights.OProfile, tempRights.OProfile);
            checkValue(ref multiRights.OReadLevel, tempRights.OReadLevel);
            checkValue(ref multiRights.OEditLevel, tempRights.OEditLevel);
            checkValue(ref multiRights.OAllowDelegate, tempRights.OAllowDelegate);
            checkValue(ref multiRights.OAllowLock, tempRights.OAllowLock);
            checkValue(ref multiRights.OAllowEmployeeRead, tempRights.OAllowEmployeeRead);
            checkValue(ref multiRights.OAllowEmployeeEdit, tempRights.OAllowEmployeeEdit);
            checkValue(ref multiRights.OAllowEmployeeImport, tempRights.OAllowEmployeeImport);
            checkValue(ref multiRights.OAllowEmployeeExport, tempRights.OAllowEmployeeExport);
            checkValue(ref multiRights.OAllowUnitAdd, tempRights.OAllowUnitAdd);
            checkValue(ref multiRights.OAllowUnitMove, tempRights.OAllowUnitMove);
            checkValue(ref multiRights.OAllowUnitDelete, tempRights.OAllowUnitDelete);
            checkValue(ref multiRights.OAllowUnitProperty, tempRights.OAllowUnitProperty);
            checkValue(ref multiRights.OAllowReportRecipient, tempRights.OAllowReportRecipient);
            checkValue(ref multiRights.OAllowStructureImport, tempRights.OAllowStructureImport);
            checkValue(ref multiRights.OAllowStructureExport, tempRights.OAllowStructureExport);
            checkValue(ref multiRights.OAllowBouncesRead, tempRights.OAllowBouncesRead);
            checkValue(ref multiRights.OAllowBouncesEdit, tempRights.OAllowBouncesEdit);
            checkValue(ref multiRights.OAllowBouncesDelete, tempRights.OAllowBouncesDelete);
            checkValue(ref multiRights.OAllowBouncesExport, tempRights.OAllowBouncesExport);
            checkValue(ref multiRights.OAllowBouncesImport, tempRights.OAllowBouncesImport);
            checkValue(ref multiRights.OAllowReInvitation, tempRights.OAllowReInvitation);
            checkValue(ref multiRights.OAllowPostNomination, tempRights.OAllowPostNomination);
            checkValue(ref multiRights.OAllowPostNominationImport, tempRights.OAllowPostNominationImport);
            checkValue(ref multiRights.FProfile, tempRights.FProfile);
            checkValue(ref multiRights.FReadLevel, tempRights.FReadLevel);
            checkValue(ref multiRights.FEditLevel, tempRights.FEditLevel);
            checkValue(ref multiRights.FSumLevel, tempRights.FSumLevel);
            checkValue(ref multiRights.FAllowCommunications, tempRights.FAllowCommunications);
            checkValue(ref multiRights.FAllowMeasures, tempRights.FAllowMeasure);
            checkValue(ref multiRights.FAllowDelete, tempRights.FAllowDelete);
            checkValue(ref multiRights.FAllowExcelExport, tempRights.FAllowExcelExport);
            checkValue(ref multiRights.FAllowReminderOnMeasure, tempRights.FAllowReminderOnMeasure);
            checkValue(ref multiRights.FAllowReminderOnUnit, tempRights.FAllowReminderOnUnit);
            checkValue(ref multiRights.FAllowUpUnitsMeasures, tempRights.FAllowUpUnitsMeasures);
            checkValue(ref multiRights.FAllowColleaguesMeasures, tempRights.FAllowColleaguesMeasures);
            checkValue(ref multiRights.FAllowThemes, tempRights.FAllowThemes);
            checkValue(ref multiRights.FAllowReminderOnThemes, tempRights.FAllowReminderOnThemes);
            checkValue(ref multiRights.FAllowReminderOnThemeUnit, tempRights.FAllowReminderOnThemeUnit);
            checkValue(ref multiRights.FAllowColleaguesThemes, tempRights.FAllowColleaguesThemes);
            checkValue(ref multiRights.FAllowOrders, tempRights.FAllowOrders);
            checkValue(ref multiRights.FAllowReminderOnOrderUnit, tempRights.FAllowReminderOnOrderUnit);
            checkValue(ref multiRights.RProfile, tempRights.RProfile);
            checkValue(ref multiRights.RReadLevel, tempRights.RReadLevel);
            checkValue(ref multiRights.DProfile, tempRights.DProfile);
            checkValue(ref multiRights.DReadLevel, tempRights.DReadLevel);
            checkValue(ref multiRights.DAllowMail, tempRights.DAllowMail);
            checkValue(ref multiRights.DAllowDownload, tempRights.DAllowDownload);
            checkValue(ref multiRights.DAllowLog, tempRights.DAllowLog);
            checkValue(ref multiRights.DAllowType1, tempRights.DAllowType1);
            checkValue(ref multiRights.DAllowType2, tempRights.DAllowType2);
            checkValue(ref multiRights.DAllowType3, tempRights.DAllowType3);
            checkValue(ref multiRights.DAllowType4, tempRights.DAllowType4);
            checkValue(ref multiRights.DAllowType5, tempRights.DAllowType5);
            checkValue(ref multiRights.AEditLevelWorkshop, tempRights.AEditLevelWorkshop);
            checkValue(ref multiRights.AReadLevelStructure, tempRights.AReadLevelStructure);
            checkValue(ref multiRights.containerID, tempRights.containerID);
            checkValue(ref multiRights.TAllowUpload, tempRights.TAllowUpload);
            checkValue(ref multiRights.TAllowDownload, tempRights.TAllowDownload);
            checkValue(ref multiRights.TAllowDeleteOwnFiles, tempRights.TAllowDeleteOwnFiles);
            checkValue(ref multiRights.TAllowDeleteAllFiles, tempRights.TAllowDeleteAllFiles);
            checkValue(ref multiRights.TAllowAccessOwnFilesWithoutPassword, tempRights.TAllowAccessOwnFilesWithoutPassword);
            checkValue(ref multiRights.TAllowAccessAllFilesWithoutPassword, tempRights.TAllowAccessAllFilesWithoutPassword);
            checkValue(ref multiRights.TAllowCreateFolder, tempRights.TAllowCreateFolder);
            checkValue(ref multiRights.TAllowDeleteOwnFiles, tempRights.TAllowDeleteOwnFiles);
            checkValue(ref multiRights.TAllowDeleteOwnFolder, tempRights.TAllowDeleteOwnFolder);
            checkValue(ref multiRights.TAllowDeleteAllFolder, tempRights.TAllowDeleteAllFolder);
            checkValue(ref multiRights.TAllowResetPassword, tempRights.TAllowResetPassword);
            checkValue(ref multiRights.TAllowTakeOwnership, tempRights.TAllowTakeOwnership);
            checkValue(ref multiRights.TAllowTaskRead, tempRights.TAllowTaskRead);
            checkValue(ref multiRights.TAllowTaskEdit, tempRights.TAllowTaskEdit);
            checkValue(ref multiRights.TAllowTaskFileUpload, tempRights.TAllowTaskFileUpload);
            checkValue(ref multiRights.TAllowTaskExport, tempRights.TAllowTaskExport);
            checkValue(ref multiRights.TAllowTaskImport, tempRights.TAllowTaskImport);
            checkValue(ref multiRights.TAllowContactRead, tempRights.TAllowContactRead);
            checkValue(ref multiRights.TAllowContactEdit, tempRights.TAllowContactEdit);
            checkValue(ref multiRights.TAllowContactExport, tempRights.TAllowContactExport);
            checkValue(ref multiRights.TAllowContactImport, tempRights.TAllowContactImport);
            checkValue(ref multiRights.AAllowAccountTransfer, tempRights.AAllowAccountTransfer);
            checkValue(ref multiRights.AAllowAccountTransfer, tempRights.AAllowAccountTransfer);
            checkValue(ref multiRights.AAllowAccountTransfer, tempRights.AAllowAccountTransfer);
            checkValue(ref multiRights.AAllowAccountManagement, tempRights.AAllowAccountManagement);
            checkValue(ref multiRights.TAllowTranslateRead, tempRights.TAllowTranslateRead);
            checkValue(ref multiRights.TAllowTranslateEdit, tempRights.TAllowTranslateEdit);
            checkValue(ref multiRights.TAllowTranslateRelease, tempRights.TAllowTranslateRealease);
            checkValue(ref multiRights.TAllowTranslateNew, tempRights.TAllowTranslateNew);
            checkValue(ref multiRights.TAllowTranslateDelete, tempRights.TAllowTranslateDelete);
            checkValue(ref multiRights.containerID, tempRights.containerID);
            checkValue(ref multiRights.roleID, tempRights.roleID);
            checkValue(ref multiRights.translateLanaguageID, tempRights.translateLanguageID);
            checkValue(ref multiRights.groupID, tempRights.groupID);
        }
        SessionObj.ActMultiRights = multiRights;

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }

    //-------------------- Threads für Mail-Versand-Vorbereitung -------------------------------
    public void doCollectMailRecipients()
    {
        // alle markierten aus der Tabelle übernehmen
        int index = 0;
        foreach (EO.Web.GridItem item in SessionObj.actUserGrid.CheckedItems)
        {
            index++;
            // in DB schreiben
            TParameterList parameterList;
            parameterList = new TParameterList();
            parameterList.addParameter("template", "string", SessionObj.actMailTemplate);
            parameterList.addParameter("userID", "string", item.Cells["userID"].Value.ToString());
            parameterList.addParameter("projectID", "string", SessionObj.Project.ProjectID);
            parameterList.addParameter("senddatetime", "datetime", SessionObj.actMailTime.ToString());
            parameterList.addParameter("job", "int", SessionObj.actMailJob.ToString());
            SqlDB dataReader = new SqlDB("");
            dataReader.execSQLwithParameter("INSERT INTO mailqueue (template, userID, projectID, senddatetime, job, version) VALUES (@template, @userID, @projectID, @senddatetime, @job, '" + SessionObj.AppVersion.ToString() + "')", parameterList);

            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, index, 0, SessionObj.actUserGrid.CheckedItems.Length, SessionObj.Project.ProjectID);
        }
        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }

    //-------------------- Threads für User-Löschung -------------------------------
    public void doDeleteUser()
    {
        // alle markierten aus der Tabelle übernehmen
        int index = 0;
        foreach (EO.Web.GridItem item in SessionObj.actUserGrid.CheckedItems)
        {
            index++;
            // in DB löschen
            TUser.delete(item.Cells["userID"].Value.ToString(), SessionObj.Project.ProjectID);

            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, index, 0, SessionObj.actUserGrid.CheckedItems.Length, SessionObj.Project.ProjectID);
        }
        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }

    //-------------------- Threads für Berechtigung-Löschung -------------------------------
    public void doDeleteRights()
    {
        // alle markierten aus der Tabelle übernehmen
        int index = 0;
        foreach (EO.Web.GridItem item in SessionObj.actUserGrid.CheckedItems)
        {
            index++;
            // in DB löschen
            TUser.deleteRight(Convert.ToInt32(item.Cells["ID"].Value), SessionObj.filterOrgTable, SessionObj.Project.ProjectID);

            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, index, 0, SessionObj.actUserGrid.CheckedItems.Length, SessionObj.Project.ProjectID);
        }
        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }

    //-------------------- Threads für User-Speicherung -------------------------------
    public void doSaveAndMailUser()
    {
        // alle markierten aus der Tabelle übernehmen
        int index = 0;

        SessionObj.ProcessMessage = "Benutzer werden gespeichert: ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, index, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);
        SessionObj.ActMultiUser.update(SessionObj.Project.ProjectID, SessionObj);

        if (SessionObj.doSendEmail)
        {
            SessionObj.ProcessMessage = "Benutzer werden gespeichert und Emails gesendet: ";
            foreach (string tempUserID in SessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, index, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);
                // Benutzer ermitteln
                TUser tempUser = new TUser(tempUserID, SessionObj.Project.ProjectID);
                // Email mit Zugangsdaten senden
                tempUser.sendMail(SessionObj.MailServerAdress, SessionObj.MailServerUser, SessionObj.MailServerPassword, true, true, SessionObj.Project.userPassword.passwordSystemValid, SessionObj.ProjectDataPath, SessionObj.Project.ProjectID, SessionObj.doEmailTemplate, SessionObj);
            }
        }

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    //-------------------- Threads für Rechte-Speicherung -------------------------------
    public void doSaveRights()
    {
        // alle markierten aus der Tabelle übernehmen
        int index = 0;

        SessionObj.ProcessMessage = "Berechtigungen werden gespeichert: ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, index, 0, SessionObj.selectedIDList.Count, SessionObj.Project.ProjectID);
        SessionObj.ActMultiRights.updateOrgRight(SessionObj.Project.ProjectID, SessionObj);

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    //-------------------- Threads für Texte-Import -------------------------------
    protected void fillTextTable(string aSheet, string tableName, ExcelFile aXlsFile, string aProjectID, ArrayList importExcelResult)
    {
        ExcelFile tempxls = aXlsFile;

        SqlDB dataReader;
        int sheetIndex = tempxls.GetSheetIndex(aSheet, false);
        if (sheetIndex > 0)
        {
            // bisherige Werte löschen
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQL("TRUNCATE TABLE " + tableName);
            importExcelResult.Add("[" + aSheet + "] vorhandene Werte gelöscht");
            // neue Werte eintragen
            tempxls.ActiveSheet = sheetIndex;
            if ((string)tempxls.GetCellValue(1, 1) == "field")
            {
                int rowIndex = 2;
                while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                {
                    int colIndex = 2;
                    while ((string)tempxls.GetCellValue(1, colIndex) != null)
                    {
                        TParameterList parameterList = new TParameterList();
                        parameterList.addParameter("field", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("text", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, colIndex)));
                        parameterList.addParameter("language", "string", Convert.ToString(tempxls.GetCellValue(1, colIndex)));
                        parameterList.addParameter("position", "int", (rowIndex-1).ToString());
                        dataReader = new SqlDB(aProjectID);
                        dataReader.execSQLwithParameter("INSERT INTO " + tableName + " (field, text, language, position) VALUES (@field, @text, @language, @position)", parameterList);
                        colIndex++;
                    }
                    rowIndex++;
                    TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                }
                importExcelResult.Add("[" + aSheet + "] neue Werte importiert");
            }
            else
            {
                // Format der Tabelle fehlerhaft
                importExcelResult.Add("[" + aSheet + "] konnte nicht importiert werden: Tabellenformat fehlerhaft");
            }
        }
    }
    public void doTextImport()
    {
        // Daten übertragen
        SessionObj.importExcelResult = new ArrayList();

        if (SessionObj.importDictionary)
        {
            SessionObj.ProcessMessage = "SIS Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("SIS", "dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }
        if (SessionObj.importOrgManagerDictionary)
        {
            SessionObj.ProcessMessage = "OrgManager Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("O.Dictionary", "orgmanager_dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }
        if (SessionObj.importResponseDictionary)
        {
            SessionObj.ProcessMessage = "Response Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("R.Dictionary", "response_dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }
        if (SessionObj.importDownloadDictionary)
        {
            SessionObj.ProcessMessage = "Download Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("D.Dictionary", "download_dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }
        if (SessionObj.importFollowUpDictionary)
        {
            SessionObj.ProcessMessage = "FollowUp Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("F.Dictionary", "followup_dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }
        if (SessionObj.importAccountManagerDictionary)
        {
            SessionObj.ProcessMessage = "AccountManager Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("A.Dictionary", "accountmanager_dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }
        if (SessionObj.importTeamspaceDictionary)
        {
            SessionObj.ProcessMessage = "Teamspace Texte werden importiert: ";
            TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
            fillTextTable("T.Dictionary", "teamspace_dictionary", SessionObj.importExcelFile, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        }

        // Excel-Datei löschen
        File.Delete(SessionObj.importFile);

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    # endregion CollectRights

    //-------------------- Threads für Struktur-Import -------------------------------
    # region StructureImport
    protected void fillStructureTable(bool aImport, int aImportType, string aSheet, string tableName, ExcelFile aXlsFile, string aProjectID, ArrayList importExcelResult)
    {
        ExcelFile tempxls = aXlsFile;

        SqlDB dataReader;
        int sheetIndex = tempxls.GetSheetIndex(aSheet, false);
        if (sheetIndex > 0)
        {
            if (aImport)
            {
                SessionObj.ProcessMessage = aSheet + " wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                if (aImportType == 0)
                {
                    // bisherige Werte löschen
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQL("TRUNCATE TABLE " + tableName);
                    importExcelResult.Add("[" + aSheet + "] vorhandene Werte gelöscht");
                }
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "value")
                {
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        int colIndex = 2;
                        while ((string)tempxls.GetCellValue(1, colIndex) != null)
                        {
                            TParameterList parameterList = new TParameterList();
                            parameterList.addParameter("value", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                            parameterList.addParameter("text", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, colIndex)));
                            parameterList.addParameter("language", "string", Convert.ToString(tempxls.GetCellValue(1, colIndex)));
                            dataReader = new SqlDB(aProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO " + tableName + " (value, text, language) VALUES (@value, @text, @language)", parameterList);
                            colIndex++;
                        }
                        rowIndex++;
                    }
                    importExcelResult.Add("[" + aSheet + "] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    importExcelResult.Add("[" + aSheet + "] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
    }
    public class TKeyValue
    {
        public string fieldID;
        public int column;
    }
    public ArrayList ImportConfig;
    public int getImportKeyColumn(string aField)
    {
        int result = 0;
        foreach (TKeyValue tempKeyValue in ImportConfig)
        {
            if (tempKeyValue.fieldID == aField)
                result = tempKeyValue.column;
        }
        return result;
    }
    public string getExcelCellValue(int aRow, int aCol, ExcelFile aExcelFile)
    {
        string result = "";
        if (aExcelFile.GetCellValue(aRow, aCol) != null)
            result = aExcelFile.GetCellValue(aRow, aCol).ToString();
        return result;
    }
    public bool EmployeeCheckStringToBeValid(string aValue, string aFieldID, bool aMadantory, int aMaxChar, string aMinSelect, string aMaxSelect, string aRegEx, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if ((aMadantory) && (aValue.Length == 0))
        {
            // Text muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }

        if ((aMaxChar > 0) && (aValue.Length > aMaxChar))
        {
            // Text zu lang
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMaxLength);
        }

        if ((aMinSelect != "") && (aMaxSelect != ""))
        {
            // Wert prüfen
            try
            {
                int tempValue = Convert.ToInt32(aValue);
                int tempMinValue = Convert.ToInt32(aMinSelect);
                int tempMaxValue = Convert.ToInt32(aMaxSelect);
                if ((tempValue < tempMinValue) || (tempValue > tempMaxValue))
                {
                    // Wert ist nicht im Bereich
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange + "(" + aValue + ")");
                }
            }
            catch
            {
                // Wert ist nicht numerisch
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange);
            }

        }
        if (aRegEx != "")
        {
            // RegEx prüfen
            if (!Regex.IsMatch(aValue, aRegEx))
            {
                // Wert entspricht nicht dem notwendigen Format
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRegEx);
            }
        }

        return result;
    }
    public bool EmployeeCheckRadioToBeValid(string aValue, string aFieldID, ArrayList aActOrgIDList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if (aValue.Length == 0)
        {
            // Eintrag muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }

        // Wert auf Gültigkeit testen
        bool valid = false;
        TEmployeeCategoriesList tempCategories = new TEmployeeCategoriesList(aFieldID, SessionObj.Project.ProjectID);
        for (int i = 0; i < tempCategories.ValueList.Count; i++)
        {
            // Test ob Kategorie verfügbar ist für akteulle OrgID
            string actRight = "allow";
            // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
            int actIndex = 0;
            while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
            {
                string tempRight = SessionObj.Project.employeeFieldsCategoriesFilterList.getRight(aFieldID, ((TEmployeeCategoriesList.TEntry)tempCategories.ValueList[i]).Value, (int)aActOrgIDList[actIndex]);
                if (tempRight != "")
                {
                    actRight = tempRight;
                    // suche beenden, da Eintrag gefunden und Einträge darüber nicht mehr relevant
                    actIndex = aActOrgIDList.Count;
                }
                else
                {
                    // zur nächsthöheren OrgID weitergehen
                    actIndex++;
                }
            }
            if (actRight == "allow")
            {
                if (aValue == ((TEmployeeCategoriesList.TEntry)tempCategories.ValueList[i]).Value.ToString())
                {
                    valid = true;
                }
            }
        }
        if (!valid)
        {
            // Eintrag ungültig
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
        }
        return result;
    }
    public bool EmployeeCheckCheckboxToBeValid(string aValue, string aFieldID, string aMinSelect, string aMaxSelect, ArrayList aActOrgIDList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        string[] aValues = aValue.Split(' ');

        // Anzahl der Werte prüfen
        if ((aMinSelect != "") && (aMaxSelect != ""))
        {
            // Wert prüfen
            try
            {
                int tempValue = aValues.GetLength(0);
                int tempMinValue = Convert.ToInt32(aMinSelect);
                int tempMaxValue = Convert.ToInt32(aMaxSelect);
                if ((tempValue < tempMinValue) || (tempValue > tempMaxValue) || (aValue.Trim() == ""))
                {
                    // Wert ist nicht im Bereich
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMaxLength);
                }
            }
            catch
            {
                // Wert ist nicht numerisch
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange + "(" + aValue + ")");
            }
        }
        // Werte auf Gültigkeit testen
        TEmployeeCategoriesList tempCategories = new TEmployeeCategoriesList(aFieldID, SessionObj.Project.ProjectID);
        foreach (string actValue in aValues)
        {
            if (actValue != "")
            {
                bool valid = false;
                for (int i = 0; i < tempCategories.ValueList.Count; i++)
                {
                    // Test ob Kategorie verfügbar ist für aktuelle OrgID
                    string actRight = "allow";
                    // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
                    int actIndex = 0;
                    while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
                    {
                        string tempRight = SessionObj.Project.employeeFieldsCategoriesFilterList.getRight(aFieldID, ((TEmployeeCategoriesList.TEntry)tempCategories.ValueList[i]).Value, (int)aActOrgIDList[actIndex]);
                        if (tempRight != "")
                        {
                            actRight = tempRight;
                            // suche beenden, da Eintrag gefunden und Einträge darüber nicht mehr relevant
                            actIndex = aActOrgIDList.Count;
                        }
                        else
                        {
                            // zur nächsthöheren OrgID weitergehen
                            actIndex++;
                        }
                    }
                    if (actRight == "allow")
                    {
                        if (actValue == ((TEmployeeCategoriesList.TEntry)tempCategories.ValueList[i]).Value.ToString())
                        {
                            valid = true;
                        }
                    }
                }
                if (!valid)
                {
                    // Eintrag ungültig
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + actValue + ")");
                }
            }
        }

        return result;
    }
    public bool EmployeeCheckDateToBeValid(string aValue, string aFieldID, bool aMadantory, string aMinSelect, string aMaxSelect, ArrayList aActOrgIDList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if ((aMadantory) && (aValue.Length == 0))
        {
            // Datum muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }

        if ((aValue != "") && (aMinSelect != "") && (aMaxSelect != ""))
        {
            // Wert auf Gültigkeit testen
            try
            {
                //DateTime tempValue = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(aValue) - 1);
                //DateTime tempValue = Convert.ToDateTime(aValue);
                DateTime tempValue;
                DateTime.TryParse(aValue, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out tempValue);
                // Datumbereich prüfen
                //DateTime tempMinValue = Convert.ToDateTime(aMinSelect);
                DateTime tempMinValue;
                DateTime.TryParse(aMinSelect, System.Globalization.CultureInfo.CreateSpecificCulture("de-DE"), System.Globalization.DateTimeStyles.None, out tempMinValue);
                //DateTime tempMaxValue = Convert.ToDateTime(aMaxSelect);
                DateTime tempMaxValue;
                DateTime.TryParse(aMaxSelect, System.Globalization.CultureInfo.CreateSpecificCulture("de-DE"), System.Globalization.DateTimeStyles.None, out tempMaxValue);
                //int tempaValue = Convert.ToInt32(aValue);
                //int tempMinValue = Convert.ToInt32(aMinSelect);
                //int tempMaxValue = Convert.ToInt32(aMaxSelect);
                if ((tempValue < tempMinValue) || (tempValue > tempMaxValue))
                //if ((tempaValue < tempMinValue) || (tempaValue > tempMaxValue))
                {
                    // Wert ist nicht im Bereich
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange + " (" + tempValue.ToShortDateString() + ")");
                }
            }
            catch
            {
                // Wert ist nicht Datum
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
            }
        }
        return result;
    }
    public bool EmployeeCheckMatrixToBeValid(string aValue, string aFieldID, string aTable, bool aMadantory, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if ((aMadantory) && ((aValue == "0") || aValue == ""))
        {
            // Matrix muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }
        else
        {
            // Test ob in Structure
            if (aValue != "0")
            {
                try
                {
                    int actValue = Convert.ToInt32(aValue);

                    SqlDB dataReader;
                    dataReader = new SqlDB("select orgID from " + aTable + " WHERE orgID='" + actValue.ToString() + "'", SessionObj.Project.ProjectID);
                    if (!dataReader.read())
                    {
                        // Wert ist keine gültige OrgID
                        result = false;
                        SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
                    }
                    dataReader.close();
                }
                catch
                {
                    // Wert ist keine Zahl
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
                }
            }
        }
        return result;
    }
    public bool EmployeeCheckOriginToBeValid(string aValue, string aFieldID, string aTable, bool aMadantory, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if (aValue == "")
        {
            // Origin muss vorhanden sein, ggf. 0
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }
        else
        {
            // Test ob in Structure
            if (aValue != "0")
            {
                try
                {
                    int actValue = Convert.ToInt32(aValue);

                    SqlDB dataReader;
                    dataReader = new SqlDB("select orgID from " + aTable + " WHERE orgID='" + actValue.ToString() + "'", SessionObj.Project.ProjectID);
                    if (!dataReader.read())
                    {
                        // Wert ist keine gültige OrgID
                        result = false;
                        SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
                    }
                    dataReader.close();
                }
                catch
                {
                    // Wert ist keine Zahl
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
                }
            }
        }
        return result;
    }
    public bool EmployeeCheckUnitManagerToBeValid(string aValue, string aFieldID, bool aMadantory, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if (aValue == "")
        {
            // Wert muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }
        else
        {
            try
            {
                int actValue = Convert.ToInt32(aValue);
            }
            catch
            {
                // Wert ist keine Zahl
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
            }
        }
        return result;
    }
    public bool StructureCheckStringToBeValid(string aValue, string aFieldID, string aTable, bool aMadantory, int aMaxChar, string aMinSelect, string aMaxSelect, string aRegEx, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel(aTable, aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if ((aMadantory) && (aValue.Length == 0))
        {
            // Text muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);

        }

        if ((aMaxChar > 0) && (aValue.Length > aMaxChar))
        {
            // Text zu lang
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMaxLength);
        }

        if ((aMinSelect != "") && (aMaxSelect != ""))
        {
            // Wert prüfen
            try
            {
                int tempValue = Convert.ToInt32(aValue);
                int tempMinValue = Convert.ToInt32(aMinSelect);
                int tempMaxValue = Convert.ToInt32(aMaxSelect);
                if ((tempValue < tempMinValue) || (tempValue > tempMaxValue))
                {
                    // Wert ist nicht im Bereich
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange + "(" + aValue + ")");
                }
            }
            catch
            {
                // Wert ist nicht numerisch
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange);
            }

        }
        if (aRegEx != "")
        {
            // RegEx prüfen
            if (!Regex.IsMatch(aValue, aRegEx))
            {
                // Wert entspricht nicht dem notwendigen Format
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRegEx);
            }
        }

        return result;
    }
    public bool StructureCheckRadioToBeValid(string aValue, string aFieldID, string aTable, ArrayList aActOrgIDList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel(aTable, aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if (aValue.Length == 0)
        {
            // Eintrag muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }

        // Wert auf Gültigkeit testen
        bool valid = false;
        TStructureCategoriesList tempCategories = new TStructureCategoriesList(aFieldID, aTable + "categories", SessionObj.Project.ProjectID);
        for (int i = 0; i < tempCategories.ValueList.Count; i++)
        {
            // Test ob Kategorie verfügbar ist für akteulle OrgID
            string actRight = "allow";
            // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
            int actIndex = 0;
            while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
            {
                string tempRight = SessionObj.Project.structureFieldsCategoriesFilterList.getRight(aFieldID, ((TStructureCategoriesList.TEntry)tempCategories.ValueList[i]).Value, (int)aActOrgIDList[actIndex]);
                if (tempRight != "")
                {
                    actRight = tempRight;
                    // suche beenden, da Eintrag gefunden und Einträge darüber nicht mehr relevant
                    actIndex = aActOrgIDList.Count;
                }
                else
                {
                    // zur nächsthöheren OrgID weitergehen
                    actIndex++;
                }
            }
            if (actRight == "allow")
            {
                if (aValue == ((TStructureCategoriesList.TEntry)tempCategories.ValueList[i]).Value.ToString())
                {
                    valid = true;
                }
            }
        }
        if (!valid)
        {
            // Eintrag ungültig
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
        }
        return result;
    }
    public bool StructureCheckCheckboxToBeValid(string aValue, string aFieldID, string aTable, string aMinSelect, string aMaxSelect, ArrayList aActOrgIDList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel(aTable, aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        string[] aValues = aValue.Split(' ');

        // Anzahl der Werte prüfen
        if ((aMinSelect != "") && (aMaxSelect != ""))
        {
            // Wert prüfen
            try
            {
                int tempValue = aValues.GetLength(0);
                int tempMinValue = Convert.ToInt32(aMinSelect);
                int tempMaxValue = Convert.ToInt32(aMaxSelect);
                if ((tempValue < tempMinValue) || (tempValue > tempMaxValue) || (aValue.Trim() == ""))
                {
                    // Wert ist nicht im Bereich
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMaxLength);
                }
            }
            catch
            {
                // Wert ist nicht numerisch
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange + "(" + aValue + ")");
            }
        }
        // Werte auf Gültigkeit testen
        TStructureCategoriesList tempCategories = new TStructureCategoriesList(aFieldID, aTable + "categories", SessionObj.Project.ProjectID);
        foreach (string actValue in aValues)
        {
            if (actValue != "")
            {
                bool valid = false;
                for (int i = 0; i < tempCategories.ValueList.Count; i++)
                {
                    // Test ob Kategorie verfügbar ist für aktuelle OrgID
                    string actRight = "allow";
                    // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
                    int actIndex = 0;
                    while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
                    {
                        string tempRight = SessionObj.Project.structureFieldsCategoriesFilterList.getRight(aFieldID, ((TStructureCategoriesList.TEntry)tempCategories.ValueList[i]).Value, (int)aActOrgIDList[actIndex]);
                        if (tempRight != "")
                        {
                            actRight = tempRight;
                            // suche beenden, da Eintrag gefunden und Einträge darüber nicht mehr relevant
                            actIndex = aActOrgIDList.Count;
                        }
                        else
                        {
                            // zur nächsthöheren OrgID weitergehen
                            actIndex++;
                        }
                    }
                    if (actRight == "allow")
                    {
                        if (actValue == ((TStructureCategoriesList.TEntry)tempCategories.ValueList[i]).Value.ToString())
                        {
                            valid = true;
                        }
                    }
                }
                if (!valid)
                {
                    // Eintrag ungültig
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + actValue + ")");
                }
            }
        }

        return result;
    }
    public bool StructureCheckDateToBeValid(string aValue, string aFieldID, string aTable, bool aMadantory, string aMinSelect, string aMaxSelect, ArrayList aActOrgIDList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel(aTable, aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if ((aMadantory) && (aValue.Length == 0))
        {
            // Datum muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }

        if ((aValue != "") && (aMinSelect != "") && (aMaxSelect != ""))
        {
            // Wert auf Gültigkeit testen
            try
            {
                //DateTime tempValue = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(aValue) - 1);
                //DateTime tempValue = Convert.ToDateTime(aValue);
                DateTime tempValue;
                DateTime.TryParse(aValue, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out tempValue);
                // Datumbereich prüfen
                //DateTime tempMinValue = Convert.ToDateTime(aMinSelect);
                DateTime tempMinValue;
                DateTime.TryParse(aMinSelect, System.Globalization.CultureInfo.CreateSpecificCulture("de-DE"), System.Globalization.DateTimeStyles.None, out tempMinValue);
                //DateTime tempMaxValue = Convert.ToDateTime(aMaxSelect);
                DateTime tempMaxValue;
                DateTime.TryParse(aMaxSelect, System.Globalization.CultureInfo.CreateSpecificCulture("de-DE"), System.Globalization.DateTimeStyles.None, out tempMaxValue);
                //int tempaValue = Convert.ToInt32(aValue);
                //int tempMinValue = Convert.ToInt32(aMinSelect);
                //int tempMaxValue = Convert.ToInt32(aMaxSelect);
                if ((tempValue < tempMinValue) || (tempValue > tempMaxValue))
                //if ((tempaValue < tempMinValue) || (tempaValue > tempMaxValue))
                {
                    // Wert ist nicht im Bereich
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorRange + " (" + tempValue.ToShortDateString() + ")");
                }
            }
            catch
            {
                // Wert ist nicht Datum
                result = false;
                SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
            }
        }
        return result;
    }
    public bool StructureCheckBackComparisonToBeValid(string aValue, string aFieldID, string aTable, string aBackTable, bool aMadantory, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel(aTable, aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if ((aMadantory) && ((aValue == "0") || aValue == ""))
        {
            // BackComparison muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }
        else
        {
            // Test ob in Structure
            if (aValue != "0")
            {
                try
                {
                    int actValue = Convert.ToInt32(aValue);

                    SqlDB dataReader;
                    dataReader = new SqlDB("select orgID from " + aBackTable + " where orgID='" + actValue.ToString() + "'", SessionObj.Project.ProjectID);
                    if (!dataReader.read())
                    {
                        // Wert ist keine gültige OrgID1
                        result = false;
                        SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
                    }
                    dataReader.close();
                }
                catch
                {
                    // Wert ist keine Zahl
                    result = false;
                    SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
                }
            }
        }
        return result;
    }
    public bool StructureCheckUserLanguageToBeValid(string aValue, string aFieldID, string aTable, ArrayList aActOrgIDList, TUserLanguage aUserLanguageList, int aExcelRow)
    {
        bool result = true;
        Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel(aTable, aFieldID, SessionObj.Language, SessionObj.Project.ProjectID);

        if (aValue.Length == 0)
        {
            // Eintrag muss vorhanden sein
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorMandatory);
        }

        // Wert auf Gültigkeit testen
        bool valid = false;
        TUserLanguage tempCategories = new TUserLanguage(SessionObj.Project.ProjectID);
        for (int i = 0; i < aUserLanguageList.ValueList.Count; i++)
        {
            // Test ob Kategorie verfügbar ist für aktuelle OrgID
            string actRight = "allow";
            // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
            int actIndex = 0;
            while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
            {
                string tempRight = SessionObj.Project.structureFieldsCategoriesFilterList.getRight(aFieldID, ((TUserLanguage.TEntry)tempCategories.ValueList[i]).Value, (int)aActOrgIDList[actIndex]);
                if (tempRight != "")
                {
                    actRight = tempRight;
                    // suche beenden, da Eintrag gefunden und Einträge darüber nicht mehr relevant
                    actIndex = aActOrgIDList.Count;
                }
                else
                {
                    // zur nächsthöheren OrgID weitergehen
                    actIndex++;
                }
            }
            if (actRight == "allow")
            {
                if (aValue == ((TUserLanguage.TEntry)tempCategories.ValueList[i]).Value.ToString())
                {
                    valid = true;
                }
            }
        }
        if (!valid)
        {
            // Eintrag ungültig
            result = false;
            SessionObj.importExcelResult.Add("[Excel-Zeile " + (aExcelRow).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + aValue + ")");
        }
        return result;
    }
    public void doStructurImport()
    {
        ExcelFile tempxls = SessionObj.importExcelFile;
        int sheetIndex;
        bool importError;
        SqlDB dataReader;

        // Daten übertragen
        SessionObj.importExcelResult = new ArrayList();

        #region Sperre und Status
        fillStructureTable(SessionObj.importStatus, SessionObj.importStatusType, "Status", "orgmanager_status", tempxls, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        fillStructureTable(SessionObj.importSperre, SessionObj.importSperreType, "Sperre", "orgmanager_lockstatus", tempxls, SessionObj.Project.ProjectID, SessionObj.importExcelResult);
        #endregion Sperre, Status und Sprache

        # region StructureBack
        sheetIndex = tempxls.GetSheetIndex("StructureBack", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructureBack)
            {
                SessionObj.ProcessMessage = "StrukturBack wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructureBackType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structureBack");
                        SessionObj.importExcelResult.Add("[StructureBack] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructureBack.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structureBack", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structureBack WHERE orgID=@orgID", parameterList);
                        }
                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        parameterList.addParameter("filter", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structureBack ON;INSERT INTO structureBack (orgId, topOrgID, displayName, displayNameShort, filter) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort, @filter);SET IDENTITY_INSERT structureBack OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[StructureBack] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[StructureBack] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[StructureBack] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion StructureBack

        # region StructureBack1
        sheetIndex = tempxls.GetSheetIndex("StructureBack1", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructureBack1)
            {
                SessionObj.ProcessMessage = "StrukturBack1 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructureBack1Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structureBack1");
                        SessionObj.importExcelResult.Add("[StructureBack1] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructureBack1.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structureBack1", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structureBack1 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        parameterList.addParameter("filter", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structureBack1 ON;INSERT INTO structureBack1 (orgId, topOrgID, displayName, displayNameShort, filter) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort, @filter);SET IDENTITY_INSERT structureBack1 OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[StructureBack1] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[StructureBack1] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[StructureBack1] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion StructureBack1

        # region StructureBack2
        sheetIndex = tempxls.GetSheetIndex("StructureBack2", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructureBack2)
            {
                SessionObj.ProcessMessage = "StrukturBack2 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructureBack2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structureBack2");
                        SessionObj.importExcelResult.Add("[StructureBack2] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructureBack2.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structureBack2", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structureBack2 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        parameterList.addParameter("filter", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structureBack2 ON;INSERT INTO structureBack2 (orgId, topOrgID, displayName, displayNameShort, filter) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort, @filter);SET IDENTITY_INSERT structureBack2 OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[StructureBack2] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[StructureBack2] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[StructureBack2] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion StructureBack1

        # region StructureBack3
        sheetIndex = tempxls.GetSheetIndex("StructureBack3", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructureBack3)
            {
                SessionObj.ProcessMessage = "StrukturBack3 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructureBack3Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structureBack3");
                        SessionObj.importExcelResult.Add("[StructureBack3] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructureBack2.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structureBack3", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structureBack3 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        parameterList.addParameter("filter", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structureBack3 ON;INSERT INTO structureBack3 (orgId, topOrgID, displayName, displayNameShort, filter) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort, @filter);SET IDENTITY_INSERT structureBack3 OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[StructureBack3] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[StructureBack3] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[StructureBack3] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion StructureBack1

        # region Structure1
        sheetIndex = tempxls.GetSheetIndex("Structure1", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructure1)
            {
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructure1Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structure1");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure1");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure1fieldsvalues");
                        SessionObj.importExcelResult.Add("[Structure1] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "Struktur1 wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure1.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structure1", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structure1 WHERE orgID=@orgID", parameterList);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure1 WHERE orgID=@orgID", parameterList);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure1fieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structure1 ON;INSERT INTO structure1 (orgId, topOrgID, displayName, displayNameShort) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort);SET IDENTITY_INSERT structure1 OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure1] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Structure1] Import abgeschlossen");

                    SessionObj.ProcessMessage = "O.Struktur1 wird importiert: ";
                    // Import-Spalten ermitteln
                    ImportConfig = new ArrayList();
                    int actCol = 6;
                    while ((tempxls.GetCellValue(1, actCol) != null) && (tempxls.GetCellValue(1, actCol).ToString() != ""))
                    {
                        TKeyValue tempKey = new TKeyValue();
                        tempKey.fieldID = tempxls.GetCellValue(1, actCol).ToString();
                        tempKey.column = actCol;
                        ImportConfig.Add(tempKey);

                        actCol++;
                    }

                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    parameterList = new TParameterList();
                    TUserLanguage userLanguageList = new TUserLanguage(SessionObj.Project.ProjectID);
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);

                    // Schleife über alle Excel-Zeilen
                    int rowIndexExcel = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)) != 0)
                    {
                        int importOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1));
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndexExcel, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)), "orgmanager_structure1", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure1 WHERE orgID=@orgID", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure1fieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        TParameterList parameterListFields = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString());
                        parameterList.addParameter("status", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 5)).ToString());
                        parameterList.addParameter("locked", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("numberSumInput", "int", "0");

                        // Schleife über Felder
                        string inserts = "";
                        bool showEmployeeCount = false;
                        bool showEmployeeCountTotal = false;
                        bool showBackComparison = false;
                        bool showLanguageUser = false;
                        importError = false;
                        int numberEmployees = 0;
                        int numberEmployeesTotal = 0;
                        int backComparison = 0;
                        // Liste der OrgIDs bis Root ermitteln
                        ArrayList actOrgIDList = new ArrayList();
                        TAdminStructure1.getOrgIDList(importOrgID, ref actOrgIDList, SessionObj.Project.ProjectID);
                        TStructureFieldsList structure1FieldList = new TStructureFieldsList(ref showEmployeeCount, ref showEmployeeCountTotal, ref showBackComparison, ref showLanguageUser, "orgmanager_structure1fields", SessionObj.Project.ProjectID);
                        foreach (TStructureFieldsList.TStructureFieldsListEntry tempEntry in structure1FieldList.structureFieldList)
                        {
                            #region Zellen einlesen
                            switch (tempEntry.fieldType)
                            {
                                case "Label":
                                case "Empty":
                                case "UnitName":
                                case "UnitNameShort":
                                    break;
                                case "Text":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Textbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Radio":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Dropdown":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Checkbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckCheckboxToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (tempValue != "")
                                            {
                                                string[] aValues = tempValue.Split(' ');

                                                // alle Werte übernehmen
                                                foreach (string actValue in aValues)
                                                {
                                                    if (inserts != "")
                                                        inserts = inserts + ",";
                                                    inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + actValue + "')";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // Leerstring speichern
                                        if (inserts != "")
                                            inserts = inserts + ",";
                                        inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                    }
                                    break;
                                case "Date":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        //string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        string tempValue = TDate.getDate(getExcelCellValue(rowIndexExcel, actCol, tempxls));
                                        //tempValue = SessionObj.fileExcel.GetCellValue(rowIndexExcel, actCol).ToString();
                                        if (StructureCheckDateToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", tempEntry.madantory, tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            //DateTime tempValueDate = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempValue) - 1);
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Datum muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "LanguageUser":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckUserLanguageToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", actOrgIDList, userLanguageList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "EmployeeCount":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployees = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "EmployeeCountTotal":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployeesTotal = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "BackComparison":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckBackComparisonToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1fields", "structureBack1", tempEntry.madantory, rowIndexExcel))
                                            backComparison = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure1fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                            #endregion Zellen einlesen
                        }

                        parameterList.addParameter("number", "int", numberEmployees.ToString());
                        parameterList.addParameter("numberTotal", "int", numberEmployeesTotal.ToString());
                        parameterList.addParameter("backComparison", "int", backComparison.ToString());
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO orgmanager_structure1 (orgID, number, numberTotal, numberSum, numberSumInput, status, backComparison, locked) VALUES (@orgID, @number, @numberTotal, @numberSum, @numberSumInput, @status, @backComparison, @locked)", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            if (inserts != "")
                            {
                                inserts = "INSERT INTO orgmanager_structure1fieldsvalues (orgID, fieldID, value) VALUES " + inserts + ";";
                                dataReader.execSQLwithParameter(inserts, parameterListFields);
                            }
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure1] Zeile " + rowIndexExcel.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Werte ungültig)");
                        }
                        rowIndexExcel++;
                    }
                    SessionObj.importExcelResult.Add("[O.Structure1] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Structure1] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // Response1
        sheetIndex = tempxls.GetSheetIndex("R.Structure1", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importRStructure1)
            {
                SessionObj.ProcessMessage = "R.Struktur1 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importRStructure1Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_structure1");
                        SessionObj.importExcelResult.Add("[R.Structure1] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "R.Struktur1 wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "response_structure1", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_structure1 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("target", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 5)).ToString());
                        parameterList.addParameter("number", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("targetSum", "int", "0");
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO response_structure1 (orgID, number, numberSum, target, targetSum) VALUES (@orgID, @number, @numberSum, @target, @targetSum)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[R.Structure1] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[R.Structure1] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[R.Structure1] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // FollowUp1
        sheetIndex = tempxls.GetSheetIndex("F.Structure1", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importFStructure1)
            {
                SessionObj.ProcessMessage = "F.Struktur1 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importFStructure1Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_structure1");
                        SessionObj.importExcelResult.Add("[F.Structure1] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "F.Struktur1 wird importiert: ";
                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "followup_structure1", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_structure1 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("mail", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());
                        parameterList.addParameter("workshop", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("livepdf", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 7)).ToString());
                        parameterList.addParameter("maxpdflevel", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 8)).ToString());
                        parameterList.addParameter("communicationfinished", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 9)).ToString());
                        //DateTime xy = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 10)) - 1);
                        DateTime xy = Convert.ToDateTime(TDate.getDate(tempxls.GetCellValue(rowIndex, 10).ToString()));
                        parameterList.addParameter("communicationfinsheddate", "datetime", xy.ToShortDateString());
                        parameterList.addParameter("participants", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 11)).ToString());
                        try
                        {
                            // Email-Gültigkeit prüfen
                            tempRegexValidator.Validate(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());

                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO followup_structure1 (orgID, mail, workshop, livepdf, maxpdflevel, communicationfinished, communicationfinsheddate, participants) VALUES (@orgID, @mail, @workshop, @livepdf, @maxpdflevel, @communicationfinished, @communicationfinsheddate, @participants)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[F.Structure1] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Email ungültig)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[F.Structure1] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[F.Structure1] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Structure1

        # region Structure2
        sheetIndex = tempxls.GetSheetIndex("Structure2", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructure2)
            {
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructure2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structure2");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure2");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure2fieldsvalues");
                        SessionObj.importExcelResult.Add("[Structure2] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "Struktur2 wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure2.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structure2", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structure2 WHERE orgID=@orgID", parameterList);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure2 WHERE orgID=@orgID", parameterList);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure2fieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structure2 ON;INSERT INTO structure2 (orgId, topOrgID, displayName, displayNameShort) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort);SET IDENTITY_INSERT structure2 OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure2] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Structure2] Import abgeschlossen");

                    // Import-Spalten ermitteln
                    ImportConfig = new ArrayList();
                    int actCol = 6;
                    while ((tempxls.GetCellValue(1, actCol) != null) && (tempxls.GetCellValue(1, actCol).ToString() != ""))
                    {
                        TKeyValue tempKey = new TKeyValue();
                        tempKey.fieldID = tempxls.GetCellValue(1, actCol).ToString();
                        tempKey.column = actCol;
                        ImportConfig.Add(tempKey);

                        actCol++;
                    }

                    SessionObj.ProcessMessage = "O.Struktur2 wird importiert: ";
                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    parameterList = new TParameterList();
                    TUserLanguage userLanguageList = new TUserLanguage(SessionObj.Project.ProjectID);
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);

                    // Schleife über alle Excel-Zeilen
                    int rowIndexExcel = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)) != 0)
                    {
                        int importOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1));
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndexExcel, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)), "orgmanager_structure2", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure2 WHERE orgID=@orgID", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure2fieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        TParameterList parameterListFields = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString());
                        parameterList.addParameter("status", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 5)).ToString());
                        parameterList.addParameter("locked", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("numberSumInput", "int", "0");

                        // Schleife über Felder
                        string inserts = "";
                        bool showEmployeeCount = false;
                        bool showEmployeeCountTotal = false;
                        bool showBackComparison = false;
                        bool showLanguageUser = false;
                        importError = false;
                        int numberEmployees = 0;
                        int numberEmployeesTotal = 0;
                        int backComparison = 0;
                        // Liste der OrgIDs bis Root ermitteln
                        ArrayList actOrgIDList = new ArrayList();
                        TAdminStructure2.getOrgIDList(importOrgID, ref actOrgIDList, SessionObj.Project.ProjectID);
                        TStructureFieldsList structure1FieldList = new TStructureFieldsList(ref showEmployeeCount, ref showEmployeeCountTotal, ref showBackComparison, ref showLanguageUser, "orgmanager_structure2fields", SessionObj.Project.ProjectID);
                        foreach (TStructureFieldsList.TStructureFieldsListEntry tempEntry in structure1FieldList.structureFieldList)
                        {
                            #region Zellen einlesen
                            switch (tempEntry.fieldType)
                            {
                                case "Label":
                                case "Empty":
                                case "UnitName":
                                case "UnitNameShort":
                                    break;
                                case "Text":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Textbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Radio":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Dropdown":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Checkbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckCheckboxToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (tempValue != "")
                                            {
                                                string[] aValues = tempValue.Split(' ');

                                                // alle Werte übernehmen
                                                foreach (string actValue in aValues)
                                                {
                                                    if (inserts != "")
                                                        inserts = inserts + ",";
                                                    inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + actValue + "')";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // Leerstring speichern
                                        if (inserts != "")
                                            inserts = inserts + ",";
                                        inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                    }
                                    break;
                                case "Date":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        //string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        string tempValue = TDate.getDate(getExcelCellValue(rowIndexExcel, actCol, tempxls));
                                        //tempValue = SessionObj.fileExcel.GetCellValue(rowIndexExcel, actCol).ToString();
                                        if (StructureCheckDateToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", tempEntry.madantory, tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            //DateTime tempValueDate = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempValue) - 1);
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Datum muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "LanguageUser":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckUserLanguageToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", actOrgIDList, userLanguageList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "EmployeeCount":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployees = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "EmployeeCountTotal":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployeesTotal = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "BackComparison":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckBackComparisonToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2fields", "structureBack2", tempEntry.madantory, rowIndexExcel))
                                            backComparison = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure2fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                            #endregion Zellen einlesen
                        }

                        parameterList.addParameter("number", "int", numberEmployees.ToString());
                        parameterList.addParameter("numberTotal", "int", numberEmployeesTotal.ToString());
                        parameterList.addParameter("backComparison", "int", backComparison.ToString());
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO orgmanager_structure2 (orgID, number, numberTotal, numberSum, numberSumInput, status, backComparison, locked) VALUES (@orgID, @number, @numberTotal, @numberSum, @numberSumInput, @status, @backComparison, @locked)", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            if (inserts != "")
                            {
                                inserts = "INSERT INTO orgmanager_structure2fieldsvalues (orgID, fieldID, value) VALUES " + inserts + ";";
                                dataReader.execSQLwithParameter(inserts, parameterListFields);
                            }
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure2] Zeile " + rowIndexExcel.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Werte ungültig)");
                        }
                        rowIndexExcel++;
                    }
                    SessionObj.importExcelResult.Add("[O.Structure2] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Structure2] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // Response2
        sheetIndex = tempxls.GetSheetIndex("R.Structure2", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importRStructure2)
            {
                SessionObj.ProcessMessage = "R.Struktur2 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importRStructure2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_structure2");
                        SessionObj.importExcelResult.Add("[R.Structure2] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "R.Struktur2 wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "response_structure2", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_structure2 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("target", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 5)).ToString());
                        parameterList.addParameter("number", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("targetSum", "int", "0");
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO response_structure2 (orgID, number, numberSum, target, targetSum) VALUES (@orgID, @number, @numberSum, @target, @targetSum)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[R.Structure2] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[R.Structure2] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[R.Structure2] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // FollowUp2
        sheetIndex = tempxls.GetSheetIndex("F.Structure2", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importFStructure2)
            {
                SessionObj.ProcessMessage = "F.Struktur2 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importFStructure2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_structure2");
                        SessionObj.importExcelResult.Add("[F.Structure2] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "F.Struktur2 wird importiert: ";
                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "followup_structure2", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_structure2 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("mail", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());
                        parameterList.addParameter("workshop", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("livepdf", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 7)).ToString());
                        parameterList.addParameter("maxpdflevel", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 8)).ToString());
                        parameterList.addParameter("communicationfinished", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 9)).ToString());
                        //DateTime xy = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 10)) - 1);
                        DateTime xy = Convert.ToDateTime(TDate.getDate(tempxls.GetCellValue(rowIndex, 10).ToString()));
                        parameterList.addParameter("communicationfinsheddate", "datetime", xy.ToShortDateString());
                        parameterList.addParameter("participants", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 11)).ToString());
                        try
                        {
                            // Email-Gültigkeit prüfen
                            tempRegexValidator.Validate(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());

                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO followup_structure2 (orgID, mail, workshop, livepdf, maxpdflevel, communicationfinished, communicationfinsheddate, participants) VALUES (@orgID, @mail, @workshop, @livepdf, @maxpdflevel, @communicationfinished, @communicationfinsheddate, @participants)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[F.Structure2] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Email ungültig)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[F.Structure2] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[F.Structure2] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Structure2

        # region Structure3
        sheetIndex = tempxls.GetSheetIndex("Structure3", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructure3)
            {
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructure3Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structure3");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure3");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure3fieldsvalues");
                        SessionObj.importExcelResult.Add("[Structure3] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "Struktur3 wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure3.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structure3", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structure3 WHERE orgID=@orgID", parameterList);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure3 WHERE orgID=@orgID", parameterList);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure3fieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structure3 ON;INSERT INTO structure3 (orgId, topOrgID, displayName, displayNameShort) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort);SET IDENTITY_INSERT structure3 OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure3] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Structure3] Import abgeschlossen");

                    SessionObj.ProcessMessage = "O.Struktur3 wird importiert: ";
                    // Import-Spalten ermitteln
                    ImportConfig = new ArrayList();
                    int actCol = 6;
                    while ((tempxls.GetCellValue(1, actCol) != null) && (tempxls.GetCellValue(1, actCol).ToString() != ""))
                    {
                        TKeyValue tempKey = new TKeyValue();
                        tempKey.fieldID = tempxls.GetCellValue(1, actCol).ToString();
                        tempKey.column = actCol;
                        ImportConfig.Add(tempKey);

                        actCol++;
                    }

                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    parameterList = new TParameterList();
                    TUserLanguage userLanguageList = new TUserLanguage(SessionObj.Project.ProjectID);
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);

                    // Schleife über alle Excel-Zeilen
                    int rowIndexExcel = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)) != 0)
                    {
                        int importOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1));
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndexExcel, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)), "orgmanager_structure3", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure3 WHERE orgID=@orgID", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure3fieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        TParameterList parameterListFields = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString());
                        parameterList.addParameter("status", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 5)).ToString());
                        parameterList.addParameter("locked", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("numberSumInput", "int", "0");

                        // Schleife über Felder
                        string inserts = "";
                        bool showEmployeeCount = false;
                        bool showEmployeeCountTotal = false;
                        bool showBackComparison = false;
                        bool showLanguageUser = false;
                        importError = false;
                        int numberEmployees = 0;
                        int numberEmployeesTotal = 0;
                        int backComparison = 0;
                        // Liste der OrgIDs bis Root ermitteln
                        ArrayList actOrgIDList = new ArrayList();
                        TAdminStructure3.getOrgIDList(importOrgID, ref actOrgIDList, SessionObj.Project.ProjectID);
                        TStructureFieldsList structure3FieldList = new TStructureFieldsList(ref showEmployeeCount, ref showEmployeeCountTotal, ref showBackComparison, ref showLanguageUser, "orgmanager_structure3fields", SessionObj.Project.ProjectID);
                        foreach (TStructureFieldsList.TStructureFieldsListEntry tempEntry in structure3FieldList.structureFieldList)
                        {
                            #region Zellen einlesen
                            switch (tempEntry.fieldType)
                            {
                                case "Label":
                                case "Empty":
                                case "UnitName":
                                case "UnitNameShort":
                                    break;
                                case "Text":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Textbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Radio":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Dropdown":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Checkbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckCheckboxToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (tempValue != "")
                                            {
                                                string[] aValues = tempValue.Split(' ');

                                                // alle Werte übernehmen
                                                foreach (string actValue in aValues)
                                                {
                                                    if (inserts != "")
                                                        inserts = inserts + ",";
                                                    inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + actValue + "')";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // Leerstring speichern
                                        if (inserts != "")
                                            inserts = inserts + ",";
                                        inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                    }
                                    break;
                                case "Date":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        //string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        string tempValue = TDate.getDate(getExcelCellValue(rowIndexExcel, actCol, tempxls));
                                        //tempValue = SessionObj.fileExcel.GetCellValue(rowIndexExcel, actCol).ToString();
                                        if (StructureCheckDateToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", tempEntry.madantory, tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            //DateTime tempValueDate = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempValue) - 1);
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Datum muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "LanguageUser":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckUserLanguageToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", actOrgIDList, userLanguageList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "EmployeeCount":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployees = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "EmployeeCountTotal":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployeesTotal = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "BackComparison":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckBackComparisonToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3fields", "structureBack3", tempEntry.madantory, rowIndexExcel))
                                            backComparison = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structure3fields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                            #endregion Zellen einlesen
                        }

                        parameterList.addParameter("number", "int", numberEmployees.ToString());
                        parameterList.addParameter("numberTotal", "int", numberEmployeesTotal.ToString());
                        parameterList.addParameter("backComparison", "int", backComparison.ToString());
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO orgmanager_structure3 (orgID, number, numberTotal, numberSum, numberSumInput, status, backComparison, locked) VALUES (@orgID, @number, @numberTotal, @numberSum, @numberSumInput, @status, @backComparison, @locked)", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            if (inserts != "")
                            {
                                inserts = "INSERT INTO orgmanager_structure3fieldsvalues (orgID, fieldID, value) VALUES " + inserts + ";";
                                dataReader.execSQLwithParameter(inserts, parameterListFields);
                            }
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure3] Zeile " + rowIndexExcel.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Werte ungültig)");
                        }
                        rowIndexExcel++;
                    }
                    SessionObj.importExcelResult.Add("[O.Structure3] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Structure3] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // Response3
        sheetIndex = tempxls.GetSheetIndex("R.Structure3", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importRStructure3)
            {
                SessionObj.ProcessMessage = "R.Struktur3 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importRStructure3Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_structure3");
                        SessionObj.importExcelResult.Add("[R.Structure3] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "R.Struktur3 wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "response_structure3", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_structure3 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("target", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 5)).ToString());
                        parameterList.addParameter("number", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("targetSum", "int", "0");
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO response_structure3 (orgID, number, numberSum, target, targetSum) VALUES (@orgID, @number, @numberSum, @target, @targetSum)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[R.Structure3] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[R.Structure3] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[R.Structure3] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // FollowUp3
        sheetIndex = tempxls.GetSheetIndex("F.Structure3", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importFStructure3)
            {
                SessionObj.ProcessMessage = "F.Struktur3 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importFStructure3Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_structure3");
                        SessionObj.importExcelResult.Add("[F.Structure3] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "F.Struktur3 wird importiert: ";
                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "followup_structure3", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_structure3 WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("mail", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());
                        parameterList.addParameter("workshop", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("livepdf", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 7)).ToString());
                        parameterList.addParameter("maxpdflevel", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 8)).ToString());
                        parameterList.addParameter("communicationfinished", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 9)).ToString());
                        //DateTime xy = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 10)) - 1);
                        DateTime xy = Convert.ToDateTime(TDate.getDate(tempxls.GetCellValue(rowIndex, 10).ToString()));
                        parameterList.addParameter("communicationfinsheddate", "datetime", xy.ToShortDateString());
                        parameterList.addParameter("participants", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 11)).ToString());
                        try
                        {
                            // Email-Gültigkeit prüfen
                            tempRegexValidator.Validate(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());

                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO followup_structure3 (orgID, mail, workshop, livepdf, maxpdflevel, communicationfinished, communicationfinsheddate, participants) VALUES (@orgID, @mail, @workshop, @livepdf, @maxpdflevel, @communicationfinished, @communicationfinsheddate, @participants)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[F.Structure3] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Email ungültig)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[F.Structure3] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[F.Structure3] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Structure3

        # region Structure
        // Struktur
        sheetIndex = tempxls.GetSheetIndex("Structure", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importStructure)
            {
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importStructureType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE structure");
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structurefieldsvalues");

                        SessionObj.importExcelResult.Add("[Structure] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "Struktur wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "structure", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM structure WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("topOrgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                        parameterList.addParameter("displayName", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        parameterList.addParameter("displayNameShort", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("SET IDENTITY_INSERT structure ON;INSERT INTO structure (orgId, topOrgID, displayName, displayNameShort) VALUES (@orgId, @topOrgID, @displayName, @displayNameShort);SET IDENTITY_INSERT structure OFF", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[Structure] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Structure] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Structure] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // OrgManager
        sheetIndex = tempxls.GetSheetIndex("O.Structure", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importOStructure)
            {
                SessionObj.ProcessMessage = "O.Struktur wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importOStructureType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structure");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_structurefieldsvalues");
                        SessionObj.importExcelResult.Add("[O.Structure] vorhandene Werte gelöscht");
                    }

                    // Import-Spalten ermitteln
                    ImportConfig = new ArrayList();
                    int actCol = 7;
                    while ((tempxls.GetCellValue(1, actCol) != null) && (tempxls.GetCellValue(1, actCol).ToString() != ""))
                    {
                        TKeyValue tempKey = new TKeyValue();
                        tempKey.fieldID = tempxls.GetCellValue(1, actCol).ToString();
                        tempKey.column = actCol;
                        ImportConfig.Add(tempKey);

                        actCol++;
                    }

                    SessionObj.ProcessMessage = "O.Struktur wird importiert: ";
                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    TParameterList parameterList = new TParameterList();
                    TUserLanguage userLanguageList = new TUserLanguage(SessionObj.Project.ProjectID);
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);

                    // Schleife über alle Excel-Zeilen
                    int rowIndexExcel = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)) != 0)
                    {
                        int importOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1));
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndexExcel, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)), "orgmanager_structure", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structure WHERE orgID=@orgID", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_structurefieldsvalues WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        TParameterList parameterListFields = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString());
                        parameterList.addParameter("status", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 5)).ToString());
                        parameterList.addParameter("locked", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("numberSumInput", "int", "0");

                        // Schleife über Felder
                        string inserts = "";
                        bool showEmployeeCount = false;
                        bool showEmployeeCountTotal = false;
                        bool showBackComparison = false;
                        bool showLanguageUser = false;
                        importError = false;
                        int numberEmployees = 0;
                        int numberEmployeesTotal = 0;
                        int backComparison = 0;
                        // Liste der OrgIDs bis Root ermitteln
                        ArrayList actOrgIDList = new ArrayList();
                        TAdminStructure.getOrgIDList(importOrgID, ref actOrgIDList, SessionObj.Project.ProjectID);
                        TStructureFieldsList structureFieldList = new TStructureFieldsList(ref showEmployeeCount, ref showEmployeeCountTotal, ref showBackComparison, ref showLanguageUser, "orgmanager_structurefields", SessionObj.Project.ProjectID);
                        foreach (TStructureFieldsList.TStructureFieldsListEntry tempEntry in structureFieldList.structureFieldList)
                        {
                            #region Zellen einlesen
                            switch (tempEntry.fieldType)
                            {
                                case "Label":
                                case "Empty":
                                case "UnitName":
                                case "UnitNameShort":
                                    break;
                                case "Text":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Textbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Radio":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Dropdown":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckRadioToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Checkbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckCheckboxToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (tempValue != "")
                                            {
                                                string[] aValues = tempValue.Split(' ');

                                                // alle Werte übernehmen
                                                foreach (string actValue in aValues)
                                                {
                                                    if (inserts != "")
                                                        inserts = inserts + ",";
                                                    inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + actValue + "')";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // Leerstring speichern
                                        if (inserts != "")
                                            inserts = inserts + ",";
                                        inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                    }
                                    break;
                                case "Date":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = TDate.getDate(getExcelCellValue(rowIndexExcel, actCol, tempxls));
                                        //tempValue = SessionObj.fileExcel.GetCellValue(rowIndexExcel, actCol).ToString();
                                        if (StructureCheckDateToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", tempEntry.madantory, tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            //DateTime tempValueDate = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempValue) - 1);
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Datum muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "LanguageUser":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckUserLanguageToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", actOrgIDList, userLanguageList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "EmployeeCount":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployees = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "EmployeeCountTotal":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckStringToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            numberEmployeesTotal = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                case "BackComparison":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (StructureCheckBackComparisonToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structurefields", "structureBack", tempEntry.madantory, rowIndexExcel))
                                            backComparison = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                            #endregion Zellen einlesen
                        }

                        parameterList.addParameter("number", "int", numberEmployees.ToString());
                        parameterList.addParameter("numberTotal", "int", numberEmployeesTotal.ToString());
                        parameterList.addParameter("backComparison", "int", backComparison.ToString());
                        //try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO orgmanager_structure (orgID, number, numberTotal, numberSum, numberSumInput, status, backComparison, locked) VALUES (@orgID, @number, @numberTotal, @numberSum, @numberSumInput, @status, @backComparison, @locked)", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            if (inserts != "")
                            {
                                inserts = "INSERT INTO orgmanager_structurefieldsvalues (orgID, fieldID, value) VALUES " + inserts + ";";
                                dataReader.execSQLwithParameter(inserts, parameterListFields);
                            }
                        }
                        //catch
                        {
                            //    SessionObj.importExcelResult.Add("[O.Structure] Zeile " + rowIndexExcel.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Werte ungültig)");
                        }
                        rowIndexExcel++;
                    }
                    SessionObj.importExcelResult.Add("[O.Structure] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[O.Structure] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // Response
        sheetIndex = tempxls.GetSheetIndex("R.Structure", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importRStructure)
            {
                SessionObj.ProcessMessage = "R.Struktur wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importRStructureType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_structure");
                        SessionObj.importExcelResult.Add("[R.Structure] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "R.Struktur wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "response_structure", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_structure WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("target", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 5)).ToString());
                        parameterList.addParameter("number", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("numberSum", "int", "0");
                        parameterList.addParameter("targetSum", "int", "0");
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO response_structure (orgID, number, numberSum, target, targetSum) VALUES (@orgID, @number, @numberSum, @target, @targetSum)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[R.Structure] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[R.Structure] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[R.Structure] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // Download
        sheetIndex = tempxls.GetSheetIndex("D.Structure", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importDStructure)
            {
                SessionObj.ProcessMessage = "D.Struktur wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importDStructureType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE download_structure");
                        SessionObj.importExcelResult.Add("[D.Structure] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "D.Struktur wird importiert: ";
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "download_structure", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM download_structure WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO download_structure (orgID) VALUES (@orgID)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[D.Structure] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[D.Structure] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[D.Structure] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        // FollowUp
        sheetIndex = tempxls.GetSheetIndex("F.Structure", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importFStructure)
            {
                SessionObj.ProcessMessage = "F.Struktur wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "OrgID")
                {
                    int selectedIndex = SessionObj.importFStructureType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_structure");
                        SessionObj.importExcelResult.Add("[F.Structure] vorhandene Werte gelöscht");
                    }

                    SessionObj.ProcessMessage = "F.Struktur wird importiert: ";
                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)) != 0)
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TAdminStructure.structureExits(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)), "followup_structure", SessionObj.Project.ProjectID)))
                        {
                            // Einheit bereits vorhanden - löschen bevor diese erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_structure WHERE orgID=@orgID", parameterList);
                        }

                        parameterList = new TParameterList();
                        parameterList.addParameter("orgID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                        parameterList.addParameter("mail", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());
                        parameterList.addParameter("workshop", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 6)).ToString());
                        parameterList.addParameter("livepdf", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 7)).ToString());
                        parameterList.addParameter("maxpdflevel", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 8)).ToString());
                        parameterList.addParameter("communicationfinished", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 9)).ToString());
                        //DateTime xy = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempxls.GetCellValue(rowIndex, 10)) - 1);
                        DateTime xy = Convert.ToDateTime(TDate.getDate(tempxls.GetCellValue(rowIndex, 10).ToString()));
                        parameterList.addParameter("communicationfinsheddate", "datetime", xy.ToShortDateString());
                        parameterList.addParameter("participants", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 11)).ToString());
                        try
                        {
                            // Email-Gültigkeit prüfen
                            tempRegexValidator.Validate(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)).Trim());

                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO followup_structure (orgID, mail, workshop, livepdf, maxpdflevel, communicationfinished, communicationfinsheddate, participants) VALUES (@orgID, @mail, @workshop, @livepdf, @maxpdflevel, @communicationfinished, @communicationfinsheddate, @participants)", parameterList);
                        }
                        catch
                        {
                            SessionObj.importExcelResult.Add("[F.Structure] Zeile " + rowIndex.ToString() + " konnte nicht importiert werden (OrgID bereits vorhanden oder Email ungültig)");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[F.Structure] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[F.Structure] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }

        # endregion Structure

        # region Employee
        sheetIndex = tempxls.GetSheetIndex("Employees", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importEmployee)
            {
                SessionObj.ProcessMessage = "Mitarbeiter werden importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "EmployeeID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importEmployeeType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_employee");
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_employeefieldsvalues");

                        SessionObj.importExcelResult.Add("[Employee] vorhandene Werte gelöscht");
                    }

                    // Import-Spalten ermitteln
                    ImportConfig = new ArrayList();
                    int actCol = 3;
                    while ((tempxls.GetCellValue(1, actCol) != null) && (tempxls.GetCellValue(1, actCol).ToString() != ""))
                    {
                        TKeyValue tempKey = new TKeyValue();
                        tempKey.fieldID = tempxls.GetCellValue(1, actCol).ToString();
                        tempKey.column = actCol;
                        ImportConfig.Add(tempKey);

                        actCol++;
                    }

                    RegexStringValidator tempRegexValidator = new RegexStringValidator("(\\w+([-+.'_]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*)?");
                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);

                    // zulässige SurveyType-Werte ermitteln
                    bool showOnline = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOnline'").intValue == 1;
                    bool showPaper = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowPaper'").intValue == 1;
                    bool showHybrid = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowHybrid'").intValue == 1;
                    bool showSMS = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowSMS'").intValue == 1;

                    // Schleife über alle Excel-Zeilen
                    int rowIndexExcel = 2;
                    while (Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)) != 0)
                    {
                        int importEmployeeID = 0;
                        try
                        {
                            importEmployeeID = Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1));
                        }
                        catch
                        {
                        }
                        int importOrgID = 0;
                        try
                        {
                            importOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 2));
                        }
                        catch
                        {
                        }
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndexExcel, 0, 0, SessionObj.Project.ProjectID);
                        if ((selectedIndex == 2) && (TEmployee.employeeExits(Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)), SessionObj.Project.ProjectID)))
                        {
                            // Teilnehmer bereits vorhanden - löschen bevor dieser erneut eingetragen werden kann
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("employeeID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_employee WHERE employeeID=@employeeID", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_employeefieldsvalues WHERE employeeID=@employeeID", parameterList);
                        }

                        parameterList = new TParameterList();
                        TParameterList parameterListFields = new TParameterList();
                        //parameterList.addParameter("employeeID", "int", importEmployeeID.ToString());
                        //parameterList.addParameter("orgID", "int", importOrgID.ToString());
                        //parameterList.addParameter("accesscode", "string", tempxls.GetCellValue(rowIndexExcel, 3).ToString());

                        // Schleife über Felder
                        string inserts = "";
                        importError = false;
                        string nameEmployee = "";
                        string firstnameEmployee = "";
                        string emailEmployee = "";
                        string mobileEmployee = "";
                        string accesscode = "";
                        int surveyType = 0;
                        int orgIDAdd1 = 0;
                        int orgIDAdd2 = 0;
                        int orgIDAdd3 = 0;
                        int supervisor = 0;
                        int supervisor1 = 0;
                        int supervisor2 = 0;
                        int supervisor3 = 0;
                        int orgIDOrigin = 0;
                        int orgIDAdd1Origin = 0;
                        int orgIDAdd2Origin = 0;
                        int orgIDAdd3Origin = 0;

                        // Liste der OrgIDs bis Root ermitteln
                        ArrayList actOrgIDList = new ArrayList();
                        TAdminStructure.getOrgIDList(importOrgID, ref actOrgIDList, SessionObj.Project.ProjectID);
                        TEmployeeFieldsList employeeFieldList = new TEmployeeFieldsList(SessionObj.Project.ProjectID);
                        foreach (TEmployeeFieldsList.TEmployeeFieldsListEntry tempEntry in employeeFieldList.employeeFieldList)
                        {
                            #region Zellen einlesen
                            switch (tempEntry.fieldType)
                            {
                                case "Label":
                                case "Empty":
                                    break;
                                case "NameEmployee":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckStringToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            nameEmployee = tempValue;
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Name muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "FirstnameEmployee":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckStringToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            firstnameEmployee = tempValue;
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Name muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "EmailEmployee":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckStringToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            emailEmployee = tempValue;
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Email muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Splate " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "MobileEmployee":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckStringToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                            mobileEmployee = tempValue;
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Email muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Splate " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Text":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckStringToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Textbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckStringToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.maxchar, tempEntry.minselect, tempEntry.maxselect, tempEntry.regex, rowIndexExcel))
                                        {
                                            parameterListFields.addParameter(tempEntry.fieldID, "string", tempValue);
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', @" + tempEntry.fieldID + ")";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (tempEntry.madantory)
                                        {
                                            if (rowIndexExcel == 2)
                                            {
                                                // Spalte mit Text muss vorhanden sein
                                                importError = true;
                                                Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                                SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                            }
                                        }
                                        else
                                        {
                                            // Leerstring speichern
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                        }
                                    }
                                    break;
                                case "Radio":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckRadioToBeValid(tempValue, tempEntry.fieldID, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_structurefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Dropdown":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckRadioToBeValid(tempValue, tempEntry.fieldID, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Radio-Zuordnung muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "Checkbox":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckCheckboxToBeValid(tempValue, tempEntry.fieldID, tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (tempValue != "")
                                            {
                                                string[] aValues = tempValue.Split(' ');

                                                // alle Werte übernehmen
                                                foreach (string actValue in aValues)
                                                {
                                                    if (inserts != "")
                                                        inserts = inserts + ",";
                                                    inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + actValue + "')";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getStructureFieldLabel("orgmanager_employeefields", tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // Leerstring speichern
                                        if (inserts != "")
                                            inserts = inserts + ",";
                                        inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '')";
                                    }
                                    break;
                                case "Date":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        //string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        string tempValue = TDate.getDate(getExcelCellValue(rowIndexExcel, actCol, tempxls));
                                        //tempValue = SessionObj.fileExcel.GetCellValue(rowIndexExcel, actCol).ToString();
                                        if (EmployeeCheckDateToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, tempEntry.minselect, tempEntry.maxselect, actOrgIDList, rowIndexExcel))
                                        {
                                            if (inserts != "")
                                                inserts = inserts + ",";
                                            //DateTime tempValueDate = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(tempValue) - 1);
                                            inserts = inserts + "('" + Convert.ToInt32(tempxls.GetCellValue(rowIndexExcel, 1)).ToString() + "', '" + tempEntry.fieldID + "', '" + tempValue + "')";
                                        }
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        if (rowIndexExcel == 2)
                                        {
                                            // Spalte mit Datum muss vorhanden sein
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("Spalte " + tempEntry.fieldID + "[" + tempLabel.title + "] fehlt");
                                        }
                                    }
                                    break;
                                case "OrgIDOrigin":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckOriginToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure", tempEntry.madantory, rowIndexExcel))
                                            orgIDOrigin = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDOrigin = 0;
                                    }
                                    break;
                                case "OrgIDOrigin1":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckOriginToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1", tempEntry.madantory, rowIndexExcel))
                                            orgIDAdd1Origin = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDAdd1Origin = 0;
                                    }
                                    break;
                                case "OrgIDOrigin2":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckOriginToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2", tempEntry.madantory, rowIndexExcel))
                                            orgIDAdd2Origin = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDAdd2Origin = 0;
                                    }
                                    break;
                                case "OrgIDOrigin3":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckOriginToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3", tempEntry.madantory, rowIndexExcel))
                                            orgIDAdd3Origin = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDAdd3Origin = 0;
                                    }
                                    break;
                                case "Matrix1":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckMatrixToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure1", tempEntry.madantory, rowIndexExcel))
                                            orgIDAdd1 = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDAdd1 = 0;
                                    }
                                    break;
                                case "Matrix2":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckMatrixToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure2", tempEntry.madantory, rowIndexExcel))
                                            orgIDAdd2 = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDAdd2 = 0;
                                    }
                                    break;
                                case "Matrix3":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckMatrixToBeValid(tempValue, tempEntry.fieldID, "orgmanager_structure3", tempEntry.madantory, rowIndexExcel))
                                            orgIDAdd3 = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        orgIDAdd3 = 0;
                                    }
                                    break;
                                case "UnitManager":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckUnitManagerToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, rowIndexExcel))
                                            supervisor = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        supervisor = 0;
                                    }
                                    break;
                                case "UnitManager1":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckUnitManagerToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, rowIndexExcel))
                                            supervisor1 = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        supervisor1 = 0;
                                    }
                                    break;
                                case "UnitManager2":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckUnitManagerToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, rowIndexExcel))
                                            supervisor2 = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        supervisor2 = 0;
                                    }
                                    break;
                                case "UnitManager3":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        if (EmployeeCheckUnitManagerToBeValid(tempValue, tempEntry.fieldID, tempEntry.madantory, rowIndexExcel))
                                            supervisor3 = Convert.ToInt32(tempValue);
                                        else
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[Zeile " + (rowIndexExcel).ToString() + "] Daten ungültig: " + tempEntry.fieldID + "[" + tempLabel.title + "]");
                                        }
                                    }
                                    else
                                    {
                                        // 0 speichern
                                        supervisor3 = 0;
                                    }
                                    break;
                                case "SurveyType":
                                    actCol = getImportKeyColumn(tempEntry.fieldID);
                                    if (actCol > 0)
                                    {
                                        string tempValue = getExcelCellValue(rowIndexExcel, actCol, tempxls);
                                        bool tempError = true;
                                        if (showOnline && tempValue == "0")
                                            tempError = false;
                                        if (showPaper && tempValue == "1")
                                            tempError = false;
                                        if (showHybrid && tempValue == "2")
                                            tempError = false;
                                        if (showSMS && tempValue == "3")
                                            tempError = false;
                                        if (tempError)
                                        {
                                            importError = true;
                                            Labeling.TFieldLabel tempLabel = Labeling.getEmployeeFieldLabel(tempEntry.fieldID, SessionObj.Language, SessionObj.Project.ProjectID);
                                            SessionObj.importExcelResult.Add("[" + Labeling.getLabel("TxtEmployeeImportConfigExcelRow", SessionObj.Language, SessionObj.Project.ProjectID) + " " + (rowIndexExcel).ToString() + "] " + tempLabel.title + ": " + tempLabel.errorInvalidValue + " (" + tempValue + ")");
                                        }
                                        else
                                            surveyType = Convert.ToInt32(tempValue);
                                    }
                                    else
                                    {
                                        surveyType = 0;
                                    }
                                    break;
                                default:
                                    break;
                            }
                            #endregion Zellen einlesen
                        }

                        //parameterList = new TParameterList();
                        parameterList.addParameter("employeeID", "int", importEmployeeID.ToString());
                        parameterList.addParameter("orgID", "int", importOrgID.ToString());
                        try
                        {
                            parameterList.addParameter("accesscode", "string", tempxls.GetCellValue(rowIndexExcel, 3).ToString());
                        }
                        catch
                        {
                            parameterList.addParameter("accesscode", "string", "");
                        }
                        parameterList.addParameter("surveytype", "string", surveyType.ToString());
                        parameterList.addParameter("name", "string", nameEmployee);
                        parameterList.addParameter("firstname", "string", firstnameEmployee);
                        parameterList.addParameter("email", "string", emailEmployee.ToString());
                        parameterList.addParameter("mobile", "string", mobileEmployee.ToString());
                        parameterList.addParameter("orgIDAdd1", "int", orgIDAdd1.ToString());
                        parameterList.addParameter("orgIDAdd2", "int", orgIDAdd2.ToString());
                        parameterList.addParameter("orgIDAdd3", "int", orgIDAdd3.ToString());
                        parameterList.addParameter("orgIDOrigin", "int", orgIDOrigin.ToString());
                        parameterList.addParameter("orgIDAdd1Origin", "int", orgIDAdd1Origin.ToString());
                        parameterList.addParameter("orgIDAdd2Origin", "int", orgIDAdd2Origin.ToString());
                        parameterList.addParameter("orgIDAdd3Origin", "int", orgIDAdd3Origin.ToString());
                        parameterList.addParameter("isSupervisor", "int", supervisor.ToString());
                        parameterList.addParameter("isSupervisor1", "int", supervisor1.ToString());
                        parameterList.addParameter("isSupervisor2", "int", supervisor2.ToString());
                        parameterList.addParameter("isSupervisor3", "int", supervisor3.ToString());
                        try
                        {
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            if (importEmployeeID == 0)
                                dataReader.execSQLwithParameter("INSERT INTO orgmanager_employee (employeeID, orgID, name, firstname, email, mobile, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, isSupervisor, isSupervisor1, isSupervisor2 isSupervisor3, accesscode, surveytype) VALUES (@employeeID, @orgID, @name, @firstname, @email, @mobile, @orgIDAdd1, @orgIDAdd2, @orgIDAdd3, @orgIDOrigin, @orgIDAdd1Origin, @orgIDAdd2Origin, @orgIDAdd3Origin, @isSupervisor, @isSupervisor1, @isSupervisor2, @isSupervisor3, @accesscode, @surveytype)", parameterList);
                            else
                                dataReader.execSQLwithParameter("SET IDENTITY_INSERT orgmanager_employee ON;INSERT INTO orgmanager_employee (employeeID, orgID, name, firstname, email, mobile, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, isSupervisor, isSupervisor1, isSupervisor2, isSupervisor3, accesscode, surveytype) VALUES" + " (@employeeID, @orgID, @name, @firstname, @email, @mobile, @orgIDAdd1, @orgIDAdd2, @orgIDAdd3, @orgIDOrigin, @orgIDAdd1Origin, @orgIDAdd2Origin, @orgIDAdd3Origin, @isSupervisor, @isSupervisor1, @isSupervisor2, @isSupervisor3, @accesscode, @surveytype);SET IDENTITY_INSERT orgmanager_employee OFF", parameterList);
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            if (inserts != "")
                            {
                                inserts = "INSERT INTO orgmanager_employeefieldsvalues (employeeID, fieldID, value) VALUES " + inserts + ";";
                                dataReader.execSQLwithParameter(inserts, parameterListFields);
                            }
                        }
                        catch (ArgumentException e)
                        {
                            SessionObj.importExcelResult.Add("[Employee] Zeile " + rowIndexExcel.ToString() + " konnte nicht importiert werden (EmployeeID bereits vorhanden oder Werte ungültig)");
                        }
                        rowIndexExcel++;
                    }
                    SessionObj.importExcelResult.Add("[Employee] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Employees] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Employee

        #region Bounces
        sheetIndex = tempxls.GetSheetIndex("Bounces", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importBounces)
            {
                SessionObj.ProcessMessage = "Bounces werden importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "Email") && ((string)tempxls.GetCellValue(1, 2) == "Mobile"))
                {
                    int selectedIndex = SessionObj.importBouncesType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("UPDATE orgmanager_employee SET isBounce='0'");

                        SessionObj.importExcelResult.Add("[Bounces] vorhandene Werte gelöscht");
                    }
                    // neue Email Bounces eintragen
                    tempxls.ActiveSheet = sheetIndex;
                    int rowIndex = 2;
                    while ((tempxls.GetCellValue(rowIndex, 1) != null) && (tempxls.GetCellValue(rowIndex, 1).ToString() != ""))
                    {
                        // Boundces importieren
                        string tempEMail = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));

                        TParameterList parameterList = new TParameterList();
                        parameterList.addParameter("email", "string", tempEMail);
                        parameterList.addParameter("isBounce", "int", "1");
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQLwithParameter("UPDATE orgmanager_employee SET isBounce=@isBounce WHERE email=@email", parameterList);

                        rowIndex++;
                    }
                    // neue Mobile Bounces eintragen
                    tempxls.ActiveSheet = sheetIndex;
                    rowIndex = 2;
                    while ((tempxls.GetCellValue(rowIndex, 2) != null) && (tempxls.GetCellValue(rowIndex, 2).ToString() != ""))
                    {
                        // Boundces importieren
                        string tempMobile = Convert.ToString(tempxls.GetCellValue(rowIndex, 2));

                        TParameterList parameterList = new TParameterList();
                        parameterList.addParameter("mobile", "string", tempMobile);
                        parameterList.addParameter("isBounce", "int", "1");
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQLwithParameter("UPDATE orgmanager_employee SET isBounce=@isBounce WHERE mobile=@mobile", parameterList);

                        rowIndex++;
                    }
                    // Anzahl und Summen der Bounces neu berechnen
                    SessionObj.ProcessMessage = "Bounces werden neu berechnet... ";
                    TAdminStructure.recalcBounces(SessionObj.Project.ProjectID);

                    SessionObj.importExcelResult.Add("[Bounces] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Bounces] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        #endregion Bounces

        // Excel-Datei löschen
        File.Delete(SessionObj.importFile);

        // Summen der Soll-Zahlen neu berechnen
        SessionObj.ProcessMessage = "Sollzahlen werden neu berechnet... (Struktur) ";
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
        TAdminStructure.recalc(SessionObj.Project.ProjectID);
        SessionObj.ProcessMessage = "Sollzahlen werden neu berechnet... (Struktur 1) ";
        TAdminStructure1.recalc(SessionObj.Project.ProjectID);
        SessionObj.ProcessMessage = "Sollzahlen werden neu berechnet... (Struktur 2) ";
        TAdminStructure2.recalc(SessionObj.Project.ProjectID);
        SessionObj.ProcessMessage = "Sollzahlen werden neu berechnet... (Struktur 3) ";
        TAdminStructure3.recalc(SessionObj.Project.ProjectID);
        SessionObj.Root = null;

        // alles neu laden
        SessionObj.reload();

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    #endregion StructureImport
    //-------------------- Threads für User-Import -------------------------------
    #region User-Import
    protected bool getBool(string aValue)
    {
        if (aValue == "1")
            return true;
        else
            return false;
    }
    protected string getXlsString(int aRow, int aCol, ExcelFile aExcelFile)
    {
        string result = "";
        try
        {
            result = aExcelFile.GetCellValue(aRow, aCol).ToString();
        }
        catch
        {
        }
        return result;
    }
    protected void setUserValue(ref string actUserValue, string newValue)
    {
        if (newValue == null)
            newValue = "";
        if (newValue != "")
            actUserValue = newValue.Trim();
    }
    protected void setUserValue(ref int actUserValue, string newValue)
    {
        if (newValue == null)
            newValue = "";
        if (newValue != "")
            actUserValue = Convert.ToInt32(newValue.Trim());
    }
    protected void setUserValue(ref bool actUserValue, string newValue)
    {
        if (newValue == null)
            newValue = "";
        if (newValue != "")
            actUserValue = getBool(newValue.Trim());
    }

    protected void handleToolRight(ref ArrayList aToolRights, string aTool, string newValue)
    {
        TUser.TToolRight tempTool = null;

        foreach (TUser.TToolRight actToolRight in aToolRights)
        {
            if (actToolRight.toolID == aTool)
            {
                // Berechtigung merken
                tempTool = actToolRight;
            }
        }
        if ((newValue == "0") && (tempTool != null))
        {
            // Berechtigung löschen
            aToolRights.Remove(tempTool);
        }
        if ((newValue == "1") && (tempTool == null))
        {
            // Berechtigung hinzufügen
            tempTool = new TUser.TToolRight();
            tempTool.toolID = aTool;
            aToolRights.Add(tempTool);
        }
    }
    public void doUserImport()
    {
        ExcelFile tempxls = SessionObj.importExcelFile;
        int sheetIndex;
        SqlDB dataReader;

        // Daten übertragen
        SessionObj.importExcelResult = new ArrayList();

        # region LoginUser
        sheetIndex = tempxls.GetSheetIndex("LoginUser", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importLoginUser)
            {
                SessionObj.ProcessMessage = "LoginUser wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);

                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "UserID")
                {
                    int selectedIndex = SessionObj.importLoginUserType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE loginuser");
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapusertool");
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid");
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid1");
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid2");
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid3");
                        dataReader.execSQL("TRUNCATE TABLE mapuserback");
                        dataReader.execSQL("TRUNCATE TABLE mapuserback1");
                        dataReader.execSQL("TRUNCATE TABLE mapuserback2");
                        dataReader.execSQL("TRUNCATE TABLE mapuserback3");
                        dataReader.execSQL("TRUNCATE TABLE mapusercontainer");
                        dataReader.execSQL("TRUNCATE TABLE mapuserrights");
                        dataReader.execSQL("TRUNCATE TABLE mapusertaskrole");
                        dataReader.execSQL("TRUNCATE TABLE mapusergroup");
                        dataReader.execSQL("TRUNCATE TABLE mapusertranslatelanguage");
                        dataReader.execSQL("TRUNCATE TABLE mapuserhidestartpage");
                        SessionObj.importExcelResult.Add("[LoginUser] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        TUser tempUser = new TUser();
                        if (selectedIndex == 2)
                        {
                            // Test ob User bereits vorhanden
                            tempUser = new TUser(Convert.ToString (tempxls.GetCellValue(rowIndex, 1)).Trim(), SessionObj.Project.ProjectID);
                        }

                        tempUser.UserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        setUserValue(ref tempUser.DisplayName, Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                        if (!tempUser.Exists)
                        {
                            string tempPassword = Convert.ToString(tempxls.GetCellValue(rowIndex, 3)).Trim();
                            if (tempPassword == "")
                            {
                                tempPassword = SessionObj.Project.userPassword.generatePassword(tempUser.UserID, SessionObj.Project.ProjectID);
                                //tempPassword = Login.generatePassword(tempUser.UserID, SessionObj.Project.userPassword.passwordLength, SessionObj.Project.userPassword.passwordComplexity, SessionObj.Project.userPassword.passwordHistory, SessionObj.Project.ProjectID);
                            }
                            tempUser.Password = TCrypt.StringVerschluesseln(tempPassword);
                        }
                        setUserValue(ref tempUser.Email, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        setUserValue(ref tempUser.Title, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        setUserValue(ref tempUser.Name, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        setUserValue(ref tempUser.Firstname, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        setUserValue(ref tempUser.Street, Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        setUserValue(ref tempUser.Zipcode, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        setUserValue(ref tempUser.City, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        setUserValue(ref tempUser.State, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        setUserValue(ref tempUser.Language, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        setUserValue(ref tempUser.Comment, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        setUserValue(ref tempUser.IsInternalUser, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        setUserValue(ref tempUser.GroupID, Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        setUserValue(ref tempUser.ForceChange, Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                        setUserValue(ref tempUser.Active, Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                        if (!tempUser.Exists)
                        {
                            tempUser.LastLogin = DateTime.Now;
                            tempUser.LastLoginError = DateTime.Now;
                            tempUser.LastLoginErrorCount = 0;
                            tempUser.SystemGeneratedPassword = true;
                            tempUser.PasswordValidDate = DateTime.Now.AddDays(SessionObj.Project.userPassword.passwordSystemValid);
                        }

                        // Tool Berechtigungen
                        handleToolRight(ref tempUser.toolRights, "orgmanager", Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                        handleToolRight(ref tempUser.toolRights, "response", Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                        handleToolRight(ref tempUser.toolRights, "download", Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));
                        handleToolRight(ref tempUser.toolRights, "teamspace", Convert.ToString(tempxls.GetCellValue(rowIndex, 21)));
                        handleToolRight(ref tempUser.toolRights, "followup", Convert.ToString(tempxls.GetCellValue(rowIndex, 22)));
                        handleToolRight(ref tempUser.toolRights, "accountmanager", Convert.ToString(tempxls.GetCellValue(rowIndex, 23)));
                        handleToolRight(ref tempUser.toolRights, "onlinereporting", Convert.ToString(tempxls.GetCellValue(rowIndex, 24)));

                        if ((selectedIndex == 2) && (tempUser.Exists))
                        {
                            // update
                            tempUser.update(SessionObj.Project.ProjectID);
                        }
                        else
                        {
                            // speichern
                            if (!tempUser.save(SessionObj.Project.ProjectID, SessionObj.Project.userPassword, false))
                                SessionObj.importExcelResult.Add("[LoginUser: " + tempUser.UserID + "] konnte nicht importiert werden: UserID mehrfach vorhanden");
                        }

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[LoginUser] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[LoginUser] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion LoginUser

        # region UserOrgID
        sheetIndex = tempxls.GetSheetIndex("UserOrgID", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserOrgID)
            {
                SessionObj.ProcessMessage = "UserOrgID wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserOrgIDType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid");
                        SessionObj.importExcelResult.Add("[MapUserOrgID] vorhandene Werte gelöscht");
                    }

                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        TParameterList parameterList;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TOrgRight tempRight = TUser.getOrgRight(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserorgid WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.OEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            setUserValue(ref tempRight.OAllowDelegate, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                            setUserValue(ref tempRight.OAllowLock, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                            setUserValue(ref tempRight.OAllowEmployeeRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                            setUserValue(ref tempRight.OAllowEmployeeEdit, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                            setUserValue(ref tempRight.OAllowEmployeeImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                            setUserValue(ref tempRight.OAllowEmployeeExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                            setUserValue(ref tempRight.OAllowUnitAdd, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                            setUserValue(ref tempRight.OAllowUnitMove, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                            setUserValue(ref tempRight.OAllowUnitDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                            setUserValue(ref tempRight.OAllowUnitProperty, Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                            setUserValue(ref tempRight.OAllowReportRecipient, Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                            setUserValue(ref tempRight.OAllowStructureImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                            setUserValue(ref tempRight.OAllowStructureExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                            setUserValue(ref tempRight.OAllowBouncesRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                            setUserValue(ref tempRight.OAllowBouncesEdit, Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));
                            setUserValue(ref tempRight.OAllowBouncesDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 21)));
                            setUserValue(ref tempRight.OAllowBouncesExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 22)));
                            setUserValue(ref tempRight.OAllowBouncesImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 23)));
                            setUserValue(ref tempRight.OAllowReInvitation, Convert.ToString(tempxls.GetCellValue(rowIndex, 24)));
                            setUserValue(ref tempRight.OAllowPostNomination, Convert.ToString(tempxls.GetCellValue(rowIndex, 25)));
                            setUserValue(ref tempRight.OAllowPostNominationImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 26)));
                            // FollowUp
                            setUserValue(ref tempRight.FProfile, getXlsString(rowIndex, 27, tempxls));
                            setUserValue(ref tempRight.FReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 28)));
                            setUserValue(ref tempRight.FEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 29)));
                            setUserValue(ref tempRight.FSumLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 30)));
                            setUserValue(ref tempRight.FAllowCommunications, Convert.ToString(tempxls.GetCellValue(rowIndex, 31)));
                            setUserValue(ref tempRight.FAllowMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 32)));
                            setUserValue(ref tempRight.FAllowDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 33)));
                            setUserValue(ref tempRight.FAllowExcelExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 34)));
                            setUserValue(ref tempRight.FAllowReminderOnMeasure, Convert.ToString(tempxls.GetCellValue(rowIndex, 35)));
                            setUserValue(ref tempRight.FAllowReminderOnUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 36)));
                            setUserValue(ref tempRight.FAllowUpUnitsMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 37)));
                            setUserValue(ref tempRight.FAllowColleaguesMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 38)));
                            setUserValue(ref tempRight.FAllowThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 39)));
                            setUserValue(ref tempRight.FAllowReminderOnThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 40)));
                            setUserValue(ref tempRight.FAllowReminderOnThemeUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 41)));
                            setUserValue(ref tempRight.FAllowTopUnitsThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 42)));
                            setUserValue(ref tempRight.FAllowColleaguesThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 43)));
                            setUserValue(ref tempRight.FAllowOrder, Convert.ToString(tempxls.GetCellValue(rowIndex, 44)));
                            setUserValue(ref tempRight.FAllowReminderOnOrderUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 45)));
                            // Response
                            setUserValue(ref tempRight.RProfile, getXlsString(rowIndex, 46, tempxls));
                            setUserValue(ref tempRight.RReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 47)));
                            // Download
                            setUserValue(ref tempRight.DProfile, getXlsString(rowIndex, 48, tempxls));
                            setUserValue(ref tempRight.DReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 49)));
                            setUserValue(ref tempRight.DAllowMail, Convert.ToString(tempxls.GetCellValue(rowIndex, 50)));
                            setUserValue(ref tempRight.DAllowDownload, Convert.ToString(tempxls.GetCellValue(rowIndex, 51)));
                            setUserValue(ref tempRight.DAllowLog, Convert.ToString(tempxls.GetCellValue(rowIndex, 52)));
                            setUserValue(ref tempRight.DAllowType1, Convert.ToString(tempxls.GetCellValue(rowIndex, 53)));
                            setUserValue(ref tempRight.DAllowType2, Convert.ToString(tempxls.GetCellValue(rowIndex, 54)));
                            setUserValue(ref tempRight.DAllowType3, Convert.ToString(tempxls.GetCellValue(rowIndex, 55)));
                            setUserValue(ref tempRight.DAllowType4, Convert.ToString(tempxls.GetCellValue(rowIndex, 56)));
                            setUserValue(ref tempRight.DAllowType5, Convert.ToString(tempxls.GetCellValue(rowIndex, 57)));
                            // AccountManager
                            setUserValue(ref tempRight.AEditLevelWorkshop, Convert.ToString(tempxls.GetCellValue(rowIndex, 58)));
                            setUserValue(ref tempRight.AReadLevelStructure, Convert.ToString(tempxls.GetCellValue(rowIndex, 59)));
                            // speichern
                            if (!TUser.saveUnitRight(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserOrgID: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserOrgID: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserOrgID] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserOrgID] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserOrgID

        # region UserOrgID1
        sheetIndex = tempxls.GetSheetIndex("UserOrgID1", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserOrgID1)
            {
                SessionObj.ProcessMessage = "UserOrgID1 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserOrgID1Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid1");
                        SessionObj.importExcelResult.Add("[MapUserOrgID1] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TOrg1Right tempRight = TUser.getOrg1Right(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserorgid1 WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.OEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            setUserValue(ref tempRight.OAllowDelegate, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                            setUserValue(ref tempRight.OAllowLock, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                            setUserValue(ref tempRight.OAllowEmployeeRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                            setUserValue(ref tempRight.OAllowEmployeeEdit, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                            setUserValue(ref tempRight.OAllowEmployeeImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                            setUserValue(ref tempRight.OAllowEmployeeExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                            setUserValue(ref tempRight.OAllowUnitAdd, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                            setUserValue(ref tempRight.OAllowUnitMove, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                            setUserValue(ref tempRight.OAllowUnitDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                            setUserValue(ref tempRight.OAllowUnitProperty, Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                            setUserValue(ref tempRight.OAllowReportRecipient, Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                            setUserValue(ref tempRight.OAllowStructureImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                            setUserValue(ref tempRight.OAllowStructureExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                            // FollowUp
                            setUserValue(ref tempRight.FProfile, getXlsString(rowIndex, 19, tempxls));
                            setUserValue(ref tempRight.FReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));
                            setUserValue(ref tempRight.FEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 21)));
                            setUserValue(ref tempRight.FSumLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 22)));
                            setUserValue(ref tempRight.FAllowCommunications, Convert.ToString(tempxls.GetCellValue(rowIndex, 23)));
                            setUserValue(ref tempRight.FAllowMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 24)));
                            setUserValue(ref tempRight.FAllowDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 25)));
                            setUserValue(ref tempRight.FAllowExcelExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 26)));
                            setUserValue(ref tempRight.FAllowReminderOnMeasure, Convert.ToString(tempxls.GetCellValue(rowIndex, 27)));
                            setUserValue(ref tempRight.FAllowReminderOnUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 28)));
                            setUserValue(ref tempRight.FAllowUpUnitsMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 29)));
                            setUserValue(ref tempRight.FAllowColleaguesMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 30)));
                            setUserValue(ref tempRight.FAllowThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 31)));
                            setUserValue(ref tempRight.FAllowReminderOnThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 32)));
                            setUserValue(ref tempRight.FAllowReminderOnThemeUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 33)));
                            setUserValue(ref tempRight.FAllowTopUnitsThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 34)));
                            setUserValue(ref tempRight.FAllowColleaguesThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 35)));
                            setUserValue(ref tempRight.FAllowOrder, Convert.ToString(tempxls.GetCellValue(rowIndex, 46)));
                            setUserValue(ref tempRight.FAllowReminderOnOrderUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 37)));
                            // Response
                            setUserValue(ref tempRight.RProfile, getXlsString(rowIndex, 38, tempxls));
                            setUserValue(ref tempRight.RReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 39)));
                            // Download
                            setUserValue(ref tempRight.DProfile, getXlsString(rowIndex, 40, tempxls));
                            setUserValue(ref tempRight.DReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 41)));
                            setUserValue(ref tempRight.DAllowMail, Convert.ToString(tempxls.GetCellValue(rowIndex, 52)));
                            setUserValue(ref tempRight.DAllowDownload, Convert.ToString(tempxls.GetCellValue(rowIndex, 53)));
                            setUserValue(ref tempRight.DAllowLog, Convert.ToString(tempxls.GetCellValue(rowIndex, 44)));
                            setUserValue(ref tempRight.DAllowType1, Convert.ToString(tempxls.GetCellValue(rowIndex, 45)));
                            setUserValue(ref tempRight.DAllowType2, Convert.ToString(tempxls.GetCellValue(rowIndex, 46)));
                            setUserValue(ref tempRight.DAllowType3, Convert.ToString(tempxls.GetCellValue(rowIndex, 47)));
                            setUserValue(ref tempRight.DAllowType4, Convert.ToString(tempxls.GetCellValue(rowIndex, 48)));
                            setUserValue(ref tempRight.DAllowType5, Convert.ToString(tempxls.GetCellValue(rowIndex, 49)));
                            // AccountManager
                            setUserValue(ref tempRight.AEditLevelWorkshop, Convert.ToString(tempxls.GetCellValue(rowIndex, 50)));
                            setUserValue(ref tempRight.AReadLevelStructure, Convert.ToString(tempxls.GetCellValue(rowIndex, 51)));
                            // speichern
                            if (!TUser.saveUnitRightOrg1(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserOrgID1: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserOrgID1: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserOrgID1] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserOrgID1] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserOrgID1

        # region UserOrgID2
        sheetIndex = tempxls.GetSheetIndex("UserOrgID2", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserOrgID2)
            {
                SessionObj.ProcessMessage = "UserOrgID2 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserOrgID2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid2");
                        SessionObj.importExcelResult.Add("[MapUserOrgID2] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TOrg2Right tempRight = TUser.getOrg2Right(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserorgid2 WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.OEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            setUserValue(ref tempRight.OAllowDelegate, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                            setUserValue(ref tempRight.OAllowLock, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                            setUserValue(ref tempRight.OAllowEmployeeRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                            setUserValue(ref tempRight.OAllowEmployeeEdit, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                            setUserValue(ref tempRight.OAllowEmployeeImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                            setUserValue(ref tempRight.OAllowEmployeeExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                            setUserValue(ref tempRight.OAllowUnitAdd, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                            setUserValue(ref tempRight.OAllowUnitMove, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                            setUserValue(ref tempRight.OAllowUnitDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                            setUserValue(ref tempRight.OAllowUnitProperty, Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                            setUserValue(ref tempRight.OAllowReportRecipient, Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                            setUserValue(ref tempRight.OAllowStructureImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                            setUserValue(ref tempRight.OAllowStructureExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                            // FollowUp
                            setUserValue(ref tempRight.FProfile, getXlsString(rowIndex, 19, tempxls));
                            setUserValue(ref tempRight.FReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));
                            setUserValue(ref tempRight.FEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 21)));
                            setUserValue(ref tempRight.FSumLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 22)));
                            setUserValue(ref tempRight.FAllowCommunications, Convert.ToString(tempxls.GetCellValue(rowIndex, 23)));
                            setUserValue(ref tempRight.FAllowMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 24)));
                            setUserValue(ref tempRight.FAllowDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 25)));
                            setUserValue(ref tempRight.FAllowExcelExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 26)));
                            setUserValue(ref tempRight.FAllowReminderOnMeasure, Convert.ToString(tempxls.GetCellValue(rowIndex, 27)));
                            setUserValue(ref tempRight.FAllowReminderOnUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 28)));
                            setUserValue(ref tempRight.FAllowUpUnitsMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 29)));
                            setUserValue(ref tempRight.FAllowColleaguesMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 30)));
                            setUserValue(ref tempRight.FAllowThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 31)));
                            setUserValue(ref tempRight.FAllowReminderOnThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 32)));
                            setUserValue(ref tempRight.FAllowReminderOnThemeUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 33)));
                            setUserValue(ref tempRight.FAllowTopUnitsThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 34)));
                            setUserValue(ref tempRight.FAllowColleaguesThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 35)));
                            setUserValue(ref tempRight.FAllowOrder, Convert.ToString(tempxls.GetCellValue(rowIndex, 46)));
                            setUserValue(ref tempRight.FAllowReminderOnOrderUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 37)));
                            // Response
                            setUserValue(ref tempRight.RProfile, getXlsString(rowIndex, 38, tempxls));
                            setUserValue(ref tempRight.RReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 39)));
                            // Download
                            setUserValue(ref tempRight.DProfile, getXlsString(rowIndex, 40, tempxls));
                            setUserValue(ref tempRight.DReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 41)));
                            setUserValue(ref tempRight.DAllowMail, Convert.ToString(tempxls.GetCellValue(rowIndex, 52)));
                            setUserValue(ref tempRight.DAllowDownload, Convert.ToString(tempxls.GetCellValue(rowIndex, 53)));
                            setUserValue(ref tempRight.DAllowLog, Convert.ToString(tempxls.GetCellValue(rowIndex, 44)));
                            setUserValue(ref tempRight.DAllowType1, Convert.ToString(tempxls.GetCellValue(rowIndex, 45)));
                            setUserValue(ref tempRight.DAllowType2, Convert.ToString(tempxls.GetCellValue(rowIndex, 46)));
                            setUserValue(ref tempRight.DAllowType3, Convert.ToString(tempxls.GetCellValue(rowIndex, 47)));
                            setUserValue(ref tempRight.DAllowType4, Convert.ToString(tempxls.GetCellValue(rowIndex, 48)));
                            setUserValue(ref tempRight.DAllowType5, Convert.ToString(tempxls.GetCellValue(rowIndex, 49)));
                            // AccountManager
                            setUserValue(ref tempRight.AEditLevelWorkshop, Convert.ToString(tempxls.GetCellValue(rowIndex, 50)));
                            setUserValue(ref tempRight.AReadLevelStructure, Convert.ToString(tempxls.GetCellValue(rowIndex, 51)));
                            // speichern
                            if (!TUser.saveUnitRightOrg2(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserOrgID2: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserOrgID2: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserOrgID2] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserOrgID2] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserOrgID2

        # region UserOrgID3
        sheetIndex = tempxls.GetSheetIndex("UserOrgID3", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserOrgID3)
            {
                SessionObj.ProcessMessage = "UserOrgID3 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserOrgID3Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserorgid3");
                        SessionObj.importExcelResult.Add("[MapUserOrgID3] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TOrg3Right tempRight = TUser.getOrg3Right(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserorgid3 WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.OEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            setUserValue(ref tempRight.OAllowDelegate, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                            setUserValue(ref tempRight.OAllowLock, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                            setUserValue(ref tempRight.OAllowEmployeeRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                            setUserValue(ref tempRight.OAllowEmployeeEdit, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                            setUserValue(ref tempRight.OAllowEmployeeImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                            setUserValue(ref tempRight.OAllowEmployeeExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                            setUserValue(ref tempRight.OAllowUnitAdd, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                            setUserValue(ref tempRight.OAllowUnitMove, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                            setUserValue(ref tempRight.OAllowUnitDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                            setUserValue(ref tempRight.OAllowUnitProperty, Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                            setUserValue(ref tempRight.OAllowReportRecipient, Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                            setUserValue(ref tempRight.OAllowStructureImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                            setUserValue(ref tempRight.OAllowStructureExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                            // FollowUp
                            setUserValue(ref tempRight.FProfile, getXlsString(rowIndex, 19, tempxls));
                            setUserValue(ref tempRight.FReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));
                            setUserValue(ref tempRight.FEditLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 21)));
                            setUserValue(ref tempRight.FSumLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 22)));
                            setUserValue(ref tempRight.FAllowCommunications, Convert.ToString(tempxls.GetCellValue(rowIndex, 23)));
                            setUserValue(ref tempRight.FAllowMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 24)));
                            setUserValue(ref tempRight.FAllowDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 25)));
                            setUserValue(ref tempRight.FAllowExcelExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 26)));
                            setUserValue(ref tempRight.FAllowReminderOnMeasure, Convert.ToString(tempxls.GetCellValue(rowIndex, 27)));
                            setUserValue(ref tempRight.FAllowReminderOnUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 28)));
                            setUserValue(ref tempRight.FAllowUpUnitsMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 29)));
                            setUserValue(ref tempRight.FAllowColleaguesMeasures, Convert.ToString(tempxls.GetCellValue(rowIndex, 30)));
                            setUserValue(ref tempRight.FAllowThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 31)));
                            setUserValue(ref tempRight.FAllowReminderOnThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 32)));
                            setUserValue(ref tempRight.FAllowReminderOnThemeUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 33)));
                            setUserValue(ref tempRight.FAllowTopUnitsThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 34)));
                            setUserValue(ref tempRight.FAllowColleaguesThemes, Convert.ToString(tempxls.GetCellValue(rowIndex, 35)));
                            setUserValue(ref tempRight.FAllowOrder, Convert.ToString(tempxls.GetCellValue(rowIndex, 46)));
                            setUserValue(ref tempRight.FAllowReminderOnOrderUnit, Convert.ToString(tempxls.GetCellValue(rowIndex, 37)));
                            // Response
                            setUserValue(ref tempRight.RProfile, getXlsString(rowIndex, 38, tempxls));
                            setUserValue(ref tempRight.RReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 39)));
                            // Download
                            setUserValue(ref tempRight.DProfile, getXlsString(rowIndex, 40, tempxls));
                            setUserValue(ref tempRight.DReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 41)));
                            setUserValue(ref tempRight.DAllowMail, Convert.ToString(tempxls.GetCellValue(rowIndex, 52)));
                            setUserValue(ref tempRight.DAllowDownload, Convert.ToString(tempxls.GetCellValue(rowIndex, 53)));
                            setUserValue(ref tempRight.DAllowLog, Convert.ToString(tempxls.GetCellValue(rowIndex, 44)));
                            setUserValue(ref tempRight.DAllowType1, Convert.ToString(tempxls.GetCellValue(rowIndex, 45)));
                            setUserValue(ref tempRight.DAllowType2, Convert.ToString(tempxls.GetCellValue(rowIndex, 46)));
                            setUserValue(ref tempRight.DAllowType3, Convert.ToString(tempxls.GetCellValue(rowIndex, 47)));
                            setUserValue(ref tempRight.DAllowType4, Convert.ToString(tempxls.GetCellValue(rowIndex, 48)));
                            setUserValue(ref tempRight.DAllowType5, Convert.ToString(tempxls.GetCellValue(rowIndex, 49)));
                            // AccountManager
                            setUserValue(ref tempRight.AEditLevelWorkshop, Convert.ToString(tempxls.GetCellValue(rowIndex, 50)));
                            setUserValue(ref tempRight.AReadLevelStructure, Convert.ToString(tempxls.GetCellValue(rowIndex, 51)));
                            // speichern
                            if (!TUser.saveUnitRightOrg3(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserOrgID3: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserOrgID3: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserOrgID3] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserOrgID3] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserOrgID3

        # region UserBack
        sheetIndex = tempxls.GetSheetIndex("UserBack", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserBack)
            {
                SessionObj.ProcessMessage = "UserBack wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserBackType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserback");
                        SessionObj.importExcelResult.Add("[MapUserBack] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TBackRight tempRight = TUser.getBackRight(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserback WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.AReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            // speichern
                            if (!TUser.saveUnitRightBack(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserBack: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserBack: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserBack] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserBack] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserBack

        # region UserBack1
        sheetIndex = tempxls.GetSheetIndex("UserBack1", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserBack1)
            {
                SessionObj.ProcessMessage = "UserBack1 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserBack1Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserback1");
                        SessionObj.importExcelResult.Add("[MapUserBack1] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TBack1Right tempRight = TUser.getBack1Right(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserback1 WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.AReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            // speichern
                            if (!TUser.saveUnitRightBack1(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserBack1: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserBack1: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserBack1] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserBack1] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserBack1

        # region UserBack2
        sheetIndex = tempxls.GetSheetIndex("UserBack2", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserBack2)
            {
                SessionObj.ProcessMessage = "UserBack2 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserBack2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserback2");
                        SessionObj.importExcelResult.Add("[MapUserBack2] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TBack2Right tempRight = TUser.getBack2Right(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserback2 WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.AReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            // speichern
                            if (!TUser.saveUnitRightBack2(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserBack2: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserBack2: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserBack2] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserBack2] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserBack2

        # region UserBack3
        sheetIndex = tempxls.GetSheetIndex("UserBack3", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserBack3)
            {
                SessionObj.ProcessMessage = "UserBack3 wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "OrgID"))
                {
                    int selectedIndex = SessionObj.importUserBack2Type;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserback3");
                        SessionObj.importExcelResult.Add("[MapUserBack3] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actOrgID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TBack3Right tempRight = TUser.getBack3Right(tempUserID, actOrgID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("orgID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserback3 WHERE userID=@userID AND orgID=@orgID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.OProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.OReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.AReadLevel, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            // speichern
                            if (!TUser.saveUnitRightBack3(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserBack3: " + tempUserID + " / " + tempRight.OrgID.ToString() + "] konnte nicht importiert werden: Kombination UserID/OrgID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserBack3: " + tempUserID + " / " + actOrgID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserBack3] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserBack3] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserBack

        # region UserContainer
        sheetIndex = tempxls.GetSheetIndex("UserContainer", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserContainer)
            {
                SessionObj.ProcessMessage = "UserContainer wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "UserID") && ((string)tempxls.GetCellValue(1, 2) == "ContainerID"))
                {
                    int selectedIndex = SessionObj.importUserContainerType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapusercontainer");
                        SessionObj.importExcelResult.Add("[MapUserContainer] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        int actContainerID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        TUser.TContainerRight tempRight = TUser.getContainerRight(tempUserID, actContainerID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("containerID", "string", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString().Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapusercontainer WHERE userID=@userID AND containerID=@containerID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.TContainerProfile, getXlsString(rowIndex, 3, tempxls));
                            setUserValue(ref tempRight.TAllowUpload, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.TAllowDownload, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            setUserValue(ref tempRight.TAllowDeleteOwnFiles, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                            setUserValue(ref tempRight.TAllowDeleteAllFiles, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                            setUserValue(ref tempRight.TAllowAccessOwnFilesWithoutPassword, Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                            setUserValue(ref tempRight.TAllowAccessAllFilesWithoutPassword, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                            setUserValue(ref tempRight.TAllowCreateFolder, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                            setUserValue(ref tempRight.TAllowDeleteOwnFolder, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                            setUserValue(ref tempRight.TAllowDeleteAllFolder, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                            setUserValue(ref tempRight.TAllowResetPassword, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                            setUserValue(ref tempRight.TAllowTakeOwnership, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));

                            // speichern
                            if (!TUser.saveContainerRight(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserContainer: " + tempUserID + " / " + tempRight.ContainerID.ToString() + "] konnte nicht importiert werden: Kombination UserID/ContainerID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserContainer: " + tempUserID + " / " + actContainerID.ToString() + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserContainer] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserContainer] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserContainer

        # region UserRights
        sheetIndex = tempxls.GetSheetIndex("UserRights", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importUserRights)
            {
                SessionObj.ProcessMessage = "UserRights wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "UserID")
                {
                    int selectedIndex = SessionObj.importUserRightsType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapuserrights");
                        SessionObj.importExcelResult.Add("[MapUserRights] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();

                        // Test ob User vorhanden
                        bool userExits = false;
                        parameterList = new TParameterList();
                        parameterList.addParameter("userID", "string", tempUserID);
                        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, SessionObj.Project.ProjectID);
                        if (dataReader.read())
                        {
                            userExits = true;
                        }
                        dataReader.close();

                        TUser.TUserRight tempRight = TUser.getUserRight(tempUserID, SessionObj.Project.ProjectID);

                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapuserrights WHERE userID=@userID", parameterList);
                        }

                        if (userExits)
                        {
                            // OrgManager
                            setUserValue(ref tempRight.TTaskProfile, getXlsString(rowIndex, 2, tempxls));
                            setUserValue(ref tempRight.TAllowTaskRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                            setUserValue(ref tempRight.TAllowTaskCreateEditDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                            setUserValue(ref tempRight.TAllowTaskFileUpload, Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                            setUserValue(ref tempRight.TAllowTaskExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                            setUserValue(ref tempRight.TAllowTaskImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                            setUserValue(ref tempRight.TContactProfile, getXlsString(rowIndex, 8, tempxls));
                            setUserValue(ref tempRight.TAllowContactRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                            setUserValue(ref tempRight.TAllowContactCreateEditDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                            setUserValue(ref tempRight.TAllowContactExport, Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                            setUserValue(ref tempRight.TAllowContactImport, Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                            setUserValue(ref tempRight.AAllowAccountTransfer, Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                            setUserValue(ref tempRight.AAllowAccountManagement, Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                            setUserValue(ref tempRight.TTranslateProfile, getXlsString(rowIndex, 15, tempxls));
                            setUserValue(ref tempRight.TAllowTranslateRead, Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                            setUserValue(ref tempRight.TAllowTranslateEdit, Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                            setUserValue(ref tempRight.TAllowTranslateRelease, Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                            setUserValue(ref tempRight.TAllowTranslateNew, Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                            setUserValue(ref tempRight.TAllowTranslateDelete, Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));

                            // speichern
                            if (!TUser.saveUserRight(tempUserID, tempRight, SessionObj.Project.ProjectID))
                                SessionObj.importExcelResult.Add("[MapUserRights: " + tempUserID + "] konnte nicht importiert werden: UserID mehrfach vorhanden");
                        }
                        else
                        {
                            // User existiert nicht, Berechtigung wird nicht importiert
                            SessionObj.importExcelResult.Add("[MapUserRights: " + tempUserID + "] konnte nicht importiert werden: UserID nicht vorhanden");
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserRights] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserRights] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserRights

        # region UserGroups
        sheetIndex = tempxls.GetSheetIndex("UserGroups", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importUserGroups)
            {
                SessionObj.ProcessMessage = "UserGroups wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "GroupID")
                {
                    int selectedIndex = SessionObj.importUserGroupsType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE loginusergroups");
                        SessionObj.importExcelResult.Add("[LoginuserGroups] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        string tempGroupID = "";
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            tempGroupID = Convert.ToString(tempxls.GetCellValue(rowIndex, 2)).Trim();
                            parameterList.addParameter("groupID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 2)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM loginusergroups WHERE groupID=@groupID", parameterList);
                        }

                        string tempRoleID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        // speichern
                        parameterList = new TParameterList();
                        parameterList.addParameter("groupID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)));
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQLwithParameter("INSERT INTO loginusergroups (groupID) VALUES (@groupID)", parameterList);

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[LoginuserGroups] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[LoginuserGroups] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserGroups

        # region UserGroup
        sheetIndex = tempxls.GetSheetIndex("UserGroup", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importUserGroup)
            {
                SessionObj.ProcessMessage = "UserGroup wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "UserID")
                {
                    int selectedIndex = SessionObj.importUserGroupType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapusergroup");
                        SessionObj.importExcelResult.Add("[MapUserGroup] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("groupID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 2)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapusergroup WHERE userID=@userID AND groupID=@groupID", parameterList);
                        }

                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        string tempRoleID = Convert.ToString(tempxls.GetCellValue(rowIndex, 2)).Trim();
                        // speichern
                        if (!TUser.saveGroupRight(tempUserID, tempRoleID, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[MapUserGroup: " + tempUserID + "/" + tempRoleID.ToString() + "] konnte nicht importiert werden: UserID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserGroup] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserGroup] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserGroup

        # region UserTaskRole
        sheetIndex = tempxls.GetSheetIndex("UserTaskRole", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importUserTaskRole)
            {
                SessionObj.ProcessMessage = "UserTaskRole wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "UserID")
                {
                    int selectedIndex = SessionObj.importUserTaskRoleType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE mapusertaskrole");
                        SessionObj.importExcelResult.Add("[MapUserTaskRole] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("userID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            parameterList.addParameter("roleID", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2)).ToString());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM mapusertaskrole WHERE userID=@userID AND roleID=@roleID", parameterList);
                        }

                        TUser.TUserRight tempRight = new TUser.TUserRight();
                        string tempUserID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        int tempRoleID = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        // speichern
                        if (!TUser.saveTaskRoleRight(tempUserID, tempRoleID, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[MapUserTaskRole: " + tempUserID + "/" + tempRoleID.ToString() + "] konnte nicht importiert werden: UserID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[MapUserTaskRole] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[MapUserTaskRole] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserTaskRole

        # region UserLanguage
        sheetIndex = tempxls.GetSheetIndex("UserLanguage", false);
        if (sheetIndex > 0)
        {
            if (SessionObj.importUserLanguage)
            {
                SessionObj.ProcessMessage = "UserLanguage wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                if (SessionObj.importUserLanguageType == 0)
                {
                    // bisherige Werte löschen
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    dataReader.execSQL("TRUNCATE TABLE loginuserlanguage");
                    SessionObj.importExcelResult.Add("[UserLanguage] vorhandene Werte gelöscht");
                }
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if (((string)tempxls.GetCellValue(1, 1) == "value") && ((string)tempxls.GetCellValue(1, 2) == "code"))
                {
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        int colIndex = 3;
                        while ((string)tempxls.GetCellValue(1, colIndex) != null)
                        {
                            TParameterList parameterList = new TParameterList();
                            parameterList.addParameter("value", "int", Convert.ToInt32(tempxls.GetCellValue(rowIndex, 1)).ToString());
                            parameterList.addParameter("code", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                            parameterList.addParameter("text", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, colIndex)));
                            parameterList.addParameter("language", "string", Convert.ToString(tempxls.GetCellValue(1, colIndex)));
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("INSERT INTO loginuserlanguage (value, code, text, language) VALUES (@value, @code, @text, @language)", parameterList);
                            colIndex++;
                        }
                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[UserLanguage] neue Werte importiert");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[UserLanguage] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion UserLanguage

        # region OrgManagerProfiles
        sheetIndex = tempxls.GetSheetIndex("OrgManagerProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManagerProfile)
            {
                SessionObj.ProcessMessage = "OrgManagerProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManagerProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles");
                        SessionObj.importExcelResult.Add("[OrgManagerProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager tempRight = new TProfiles.TProfileSetOrgManager();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.OEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.OAllowDelegate = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.OAllowLock = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.OAllowEmployeeRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.OAllowEmployeeEdit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.OAllowEmployeeImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.OAllowEmployeeExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.OAllowUnitAdd = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.OAllowUnitMove = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.OAllowUnitDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.OAllowUnitProperty = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.OAllowReportRecipient = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.OAllowStructureImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        tempRight.OAllowStructureExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                        tempRight.OAllowBouncesRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                        tempRight.OAllowBouncesEdit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                        tempRight.OAllowBouncesDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                        tempRight.OAllowBouncesExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 20)));
                        tempRight.OAllowBouncesImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 21)));
                        tempRight.OAllowReInvitation = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 22)));
                        tempRight.OAllowPostNomination = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 23)));
                        tempRight.OAllowPostNominationImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 24)));
                        if (!TProfiles.saveProfileOrgManager(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManagerProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManagerProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManagerProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManagerProfiles

        # region OrgManagerBackProfiles
        sheetIndex = tempxls.GetSheetIndex("OrgManagerBackProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManagerBackProfile)
            {
                SessionObj.ProcessMessage = "OrgManagerBackProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager1BackProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profilesback");
                        SessionObj.importExcelResult.Add("[OrgManagerBackProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profilesback WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManagerBack tempRight = new TProfiles.TProfileSetOrgManagerBack();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileOrgManagerBack(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManagerBackProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManagerBackProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManagerBackProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManagerBackProfiles

        # region OrgManager1Profiles
        sheetIndex = tempxls.GetSheetIndex("OrgManager1Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManager1Profile)
            {
                SessionObj.ProcessMessage = "OrgManager1Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager1ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles1");
                        SessionObj.importExcelResult.Add("[OrgManager1Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles1 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager1 tempRight = new TProfiles.TProfileSetOrgManager1();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.OEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.OAllowLock = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.OAllowEmployeeRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.OAllowEmployeeEdit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.OAllowEmployeeImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.OAllowEmployeeExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.OAllowUnitAdd = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.OAllowUnitMove = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.OAllowUnitDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.OAllowUnitProperty = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.OAllowReportRecipient = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.OAllowStructureImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.OAllowStructureExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        if (!TProfiles.saveProfileOrgManager1(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManager1Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManager1Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManager1Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManager1Profiles

        # region OrgManager1BackProfiles
        sheetIndex = tempxls.GetSheetIndex("OrgManager1BackProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManager1BackProfile)
            {
                SessionObj.ProcessMessage = "OrgManager1BackProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager1BackProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles1back");
                        SessionObj.importExcelResult.Add("[OrgManager1BackProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles1back WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager1Back tempRight = new TProfiles.TProfileSetOrgManager1Back();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileOrgManager1Back(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManager1BackProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManager1BackProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManager1BackProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManager1BackProfiles

        # region OrgManager2Profiles
        sheetIndex = tempxls.GetSheetIndex("OrgManager2Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManager2Profile)
            {
                SessionObj.ProcessMessage = "OrgManager2Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager2ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles2");
                        SessionObj.importExcelResult.Add("[OrgManager2Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles2 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager2 tempRight = new TProfiles.TProfileSetOrgManager2();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.OEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.OAllowLock = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.OAllowEmployeeRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.OAllowEmployeeEdit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.OAllowEmployeeImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.OAllowEmployeeExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.OAllowUnitAdd = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.OAllowUnitMove = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.OAllowUnitDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.OAllowUnitProperty = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.OAllowReportRecipient = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.OAllowStructureImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.OAllowStructureExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        if (!TProfiles.saveProfileOrgManager2(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManager2Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManager2Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManager2Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManager2Profiles

        # region OrgManager2BackProfiles
        sheetIndex = tempxls.GetSheetIndex("OrgManager2BackProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManager2BackProfile)
            {
                SessionObj.ProcessMessage = "OrgManager2BackProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager2BackProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles2back");
                        SessionObj.importExcelResult.Add("[OrgManager2BackProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles2back WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager2Back tempRight = new TProfiles.TProfileSetOrgManager2Back();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileOrgManager2Back(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManager1BackProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManager2BackProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManager2BackProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManager2BackProfiles

        # region OrgManager3Profiles
        sheetIndex = tempxls.GetSheetIndex("OrgManager3Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManager3Profile)
            {
                SessionObj.ProcessMessage = "OrgManager3Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager2ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles3");
                        SessionObj.importExcelResult.Add("[OrgManager3Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles3 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager3 tempRight = new TProfiles.TProfileSetOrgManager3();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.OEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.OAllowLock = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.OAllowEmployeeRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.OAllowEmployeeEdit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.OAllowEmployeeImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.OAllowEmployeeExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.OAllowUnitAdd = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.OAllowUnitMove = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.OAllowUnitDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.OAllowUnitProperty = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.OAllowReportRecipient = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.OAllowStructureImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.OAllowStructureExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        if (!TProfiles.saveProfileOrgManager3(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManager3Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManager3Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManager3Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManager3Profiles

        # region OrgManager3BackProfiles
        sheetIndex = tempxls.GetSheetIndex("OrgManager3BackProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importOrgManager3BackProfile)
            {
                SessionObj.ProcessMessage = "OrgManager3BackProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importOrgManager3BackProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE orgmanager_profiles3back");
                        SessionObj.importExcelResult.Add("[OrgManager3BackProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM orgmanager_profiles3back WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetOrgManager3Back tempRight = new TProfiles.TProfileSetOrgManager3Back();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.OReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileOrgManager3Back(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[OrgManager3BackProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[OrgManager3BackProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[OrgManager3BackProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion OrgManager2BackProfiles

        # region ResponseProfiles
        sheetIndex = tempxls.GetSheetIndex("ResponseProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importResponseProfile)
            {
                SessionObj.ProcessMessage = "ResponseProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importResponseProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_profiles");
                        SessionObj.importExcelResult.Add("[ResponseProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_profiles WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetResponse tempRight = new TProfiles.TProfileSetResponse();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.RReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileResponse(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[ResponseProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[ResponseProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[ResponseProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion ResponseProfiles

        # region Response1Profiles
        sheetIndex = tempxls.GetSheetIndex("Response1Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importResponse1Profile)
            {
                SessionObj.ProcessMessage = "Response1Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importResponse1ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_profiles1");
                        SessionObj.importExcelResult.Add("[Response1Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_profiles1 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetResponse1 tempRight = new TProfiles.TProfileSetResponse1();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.RReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileResponse1(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[Response1Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Response1Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Response1Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Response1Profiles

        # region Response2Profiles
        sheetIndex = tempxls.GetSheetIndex("Response2Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importResponse2Profile)
            {
                SessionObj.ProcessMessage = "Response2Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importResponse2ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_profiles2");
                        SessionObj.importExcelResult.Add("[Response2Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_profiles2 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetResponse2 tempRight = new TProfiles.TProfileSetResponse2();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.RReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileResponse2(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[Response2Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Response2Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Response2Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Response2Profiles

        # region Response3Profiles
        sheetIndex = tempxls.GetSheetIndex("Response3Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importResponse3Profile)
            {
                SessionObj.ProcessMessage = "Response3Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importResponse3ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE response_profiles3");
                        SessionObj.importExcelResult.Add("[Response3Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM response_profiles3 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetResponse3 tempRight = new TProfiles.TProfileSetResponse3();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.RReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        if (!TProfiles.saveProfileResponse3(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[Response3Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Response3Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Response3Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Response3Profiles

        # region DownloadProfiles
        sheetIndex = tempxls.GetSheetIndex("DownloadProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importDownloadProfile)
            {
                SessionObj.ProcessMessage = "DownloadProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importDownloadProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE download_profiles");
                        SessionObj.importExcelResult.Add("[DownloadProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM download_profiles WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetDownload tempRight = new TProfiles.TProfileSetDownload();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.DReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.DAllowMail = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.DAllowDownload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.DAllowLog = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.DAllowType1 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.DAllowType2 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.DAllowType3 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.DAllowType4 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.DAllowType5 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        if (!TProfiles.saveProfileDownload(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[DownloadProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[DownloadProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[DownloadProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region Download1Profiles
        sheetIndex = tempxls.GetSheetIndex("Download1Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importDownload1Profile)
            {
                SessionObj.ProcessMessage = "Download1Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importDownload1ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE download_profiles1");
                        SessionObj.importExcelResult.Add("[Download1Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM download_profiles1 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetDownload1 tempRight = new TProfiles.TProfileSetDownload1();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.DReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.DAllowMail = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.DAllowDownload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.DAllowLog = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.DAllowType1 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.DAllowType2 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.DAllowType3 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.DAllowType4 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.DAllowType5 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        if (!TProfiles.saveProfileDownload1(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[Download1Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Download1Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Download1Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Download1Profiles

        # region Download2Profiles
        sheetIndex = tempxls.GetSheetIndex("Download2Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importDownload2Profile)
            {
                SessionObj.ProcessMessage = "Download2Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importDownload2ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE download_profiles2");
                        SessionObj.importExcelResult.Add("[Download2Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM download_profiles2 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetDownload2 tempRight = new TProfiles.TProfileSetDownload2();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.DReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.DAllowMail = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.DAllowDownload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.DAllowLog = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.DAllowType1 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.DAllowType2 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.DAllowType3 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.DAllowType4 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.DAllowType5 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        if (!TProfiles.saveProfileDownload2(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[Download2Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Download2Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Download2Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Download2Profiles

        # region Download3Profiles
        sheetIndex = tempxls.GetSheetIndex("Download3Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importDownload3Profile)
            {
                SessionObj.ProcessMessage = "Download3Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importDownload3ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE download_profiles3");
                        SessionObj.importExcelResult.Add("[Download3Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM download_profiles3 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetDownload3 tempRight = new TProfiles.TProfileSetDownload3();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.DReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.DAllowMail = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.DAllowDownload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.DAllowLog = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.DAllowType1 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.DAllowType2 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.DAllowType3 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.DAllowType4 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.DAllowType5 = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        if (!TProfiles.saveProfileDownload3(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[Download3Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Download3Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[Download3Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion Download3Profiles

        # region FollowUpProfiles
        sheetIndex = tempxls.GetSheetIndex("FollowUpProfiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importFollowUpProfile)
            {
                SessionObj.ProcessMessage = "FollowUpProfiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importFollowUpProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_profiles");
                        SessionObj.importExcelResult.Add("[FollowUpProfiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_profiles WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetFollowUp tempRight = new TProfiles.TProfileSetFollowUp();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.FReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.FEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.FSumLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 4));
                        tempRight.FAllowCommunications = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.FAllowMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.FAllowDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.FAllowExcelExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.FAllowReminderOnMeasure = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.FAllowReminderOnUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.FAllowUpUnitsMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.FAllowColleaguesMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.FAllowThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.FAllowReminderOnTheme = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.FAllowReminderOnThemeUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        tempRight.FAllowTopUnitsThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                        tempRight.FAllowColleaguesThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                        tempRight.FAllowOrders = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                        tempRight.FAllowReminderOnOrderUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                        if (!TProfiles.saveProfileFollowUp(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[FollowUpProfiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[FollowUpProfiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[FollowUpProfiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region FollowUp1Profiles
        sheetIndex = tempxls.GetSheetIndex("FollowUp1Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importFollowUp1Profile)
            {
                SessionObj.ProcessMessage = "FollowUp1Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importFollowUp1ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_profiles1");
                        SessionObj.importExcelResult.Add("[FollowUp1Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_profiles1 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetFollowUp1 tempRight = new TProfiles.TProfileSetFollowUp1();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.FReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.FEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.FSumLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 4));
                        tempRight.FAllowCommunications = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.FAllowMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.FAllowDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.FAllowExcelExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.FAllowReminderOnMeasure = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.FAllowReminderOnUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.FAllowUpUnitsMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.FAllowColleaguesMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.FAllowThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.FAllowReminderOnTheme = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.FAllowReminderOnThemeUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        tempRight.FAllowTopUnitsThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                        tempRight.FAllowColleaguesThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                        tempRight.FAllowOrders = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                        tempRight.FAllowReminderOnOrderUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                        if (!TProfiles.saveProfileFollowUp1(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[FollowUp1Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[FollowUp1Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[FollowUp1Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region FollowUp2Profiles
        sheetIndex = tempxls.GetSheetIndex("FollowUp2Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importFollowUp2Profile)
            {
                SessionObj.ProcessMessage = "FollowUp2Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importFollowUp2ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_profiles2");
                        SessionObj.importExcelResult.Add("[FollowUp2Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_profiles2 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetFollowUp2 tempRight = new TProfiles.TProfileSetFollowUp2();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.FReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.FEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.FSumLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 4));
                        tempRight.FAllowCommunications = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.FAllowMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.FAllowDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.FAllowExcelExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.FAllowReminderOnMeasure = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.FAllowReminderOnUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.FAllowUpUnitsMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.FAllowColleaguesMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.FAllowThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.FAllowReminderOnTheme = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.FAllowReminderOnThemeUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        tempRight.FAllowTopUnitsThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                        tempRight.FAllowColleaguesThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                        tempRight.FAllowOrders = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                        tempRight.FAllowReminderOnOrderUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                        if (!TProfiles.saveProfileFollowUp2(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[FollowUp2Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[FollowUp2Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[FollowUp2Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region FollowUp3Profiles
        sheetIndex = tempxls.GetSheetIndex("FollowUp3Profiles", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importFollowUp3Profile)
            {
                SessionObj.ProcessMessage = "FollowUp3Profiles wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importFollowUp3ProfileType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE followup_profiles3");
                        SessionObj.importExcelResult.Add("[FollowUp3Profiles] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM followup_profiles3 WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetFollowUp3 tempRight = new TProfiles.TProfileSetFollowUp3();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.FReadLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 2));
                        tempRight.FEditLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 3));
                        tempRight.FSumLevel = Convert.ToInt32(tempxls.GetCellValue(rowIndex, 4));
                        tempRight.FAllowCommunications = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.FAllowMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.FAllowDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.FAllowExcelExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.FAllowReminderOnMeasure = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.FAllowReminderOnUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.FAllowUpUnitsMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.FAllowColleaguesMeasures = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        tempRight.FAllowThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 13)));
                        tempRight.FAllowReminderOnTheme = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 14)));
                        tempRight.FAllowReminderOnThemeUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 15)));
                        tempRight.FAllowTopUnitsThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 16)));
                        tempRight.FAllowColleaguesThemes = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 17)));
                        tempRight.FAllowOrders = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 18)));
                        tempRight.FAllowReminderOnOrderUnit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 19)));
                        if (!TProfiles.saveProfileFollowUp3(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[FollowUp3Profiles: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[FollowUp3Profiles] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[FollowUp3Profiles] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region TeamspaceProfilesTasks
        sheetIndex = tempxls.GetSheetIndex("TeamspaceProfilesTasks", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importTeamspaceProfileTasks)
            {
                SessionObj.ProcessMessage = "TeamspaceProfilesTasks wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importTeamspaceProfileTasksType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE teamspace_profilestasks");
                        SessionObj.importExcelResult.Add("[TeamspaceProfilesTasks] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM teamspace_profilestasks WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetProjectplan tempRight = new TProfiles.TProfileSetProjectplan();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.AllowTaskRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                        tempRight.AllowTaskCreateEditDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.AllowTaskFileUpload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.AllowTaskExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.AllowTaskImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        if (!TProfiles.saveProfileTeamspaceTasks(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[TeamspaceProfilesTasks: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesTasks] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesTasks] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region TeamspaceProfilesContacts
        sheetIndex = tempxls.GetSheetIndex("TeamspaceProfilesContacts", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importTeamspaceProfileContacts)
            {
                SessionObj.ProcessMessage = "TeamspaceProfilesContacts wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importTeamspaceProfileContactsType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE teamspace_profilescontacts");
                        SessionObj.importExcelResult.Add("[TeamspaceProfilesContacts] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM teasmspace_profilescontacts WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetContacts tempRight = new TProfiles.TProfileSetContacts();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.AllowContactRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                        tempRight.AllowContactCreateEditDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.AllowContactImport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.AllowContactExport = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        if (!TProfiles.saveProfileTeamspaceContacts(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[TeamspaceProfilesContacts: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesContacts] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesContacts] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion DownloadProfiles

        # region TeamspaceProfilesContainer
        sheetIndex = tempxls.GetSheetIndex("TeamspaceProfilesContainer", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importTeamspaceProfileContainer)
            {
                SessionObj.ProcessMessage = "TeamspaceProfilesContainer wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importTeamspaceProfileContainerType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE teamspace_profilescontainer");
                        SessionObj.importExcelResult.Add("[TeamspaceProfilesContainer] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM teamspace_profilescontainer WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetContainer tempRight = new TProfiles.TProfileSetContainer();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.allowUpload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                        tempRight.allowDownload = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.allowDeleteOwnFiles = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.allowDeleteAllFiles = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.allowAccessOwnFilesWithoutPassword = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        tempRight.allowAccessAllFilesWithoutPassword = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 7)));
                        tempRight.allowCreateFolder = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 8)));
                        tempRight.allowDeleteOwnFolder = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 9)));
                        tempRight.allowDeleteAllFolder = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 10)));
                        tempRight.allowResetPassword = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 11)));
                        tempRight.allowTakeOwnership = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 12)));
                        if (!TProfiles.saveProfileTeamspaceContainer(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[TeamspaceProfilesContainer: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesContainer] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesContainer] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion TeamspaceProfilesContainer

        # region TeamspaceProfilesTranslate
        sheetIndex = tempxls.GetSheetIndex("TeamspaceProfilesTranslate", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importTeamspaceProfileTranslate)
            {
                SessionObj.ProcessMessage = "TeamspaceProfilesTranslate wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "ProfileID")
                {
                    int selectedIndex = SessionObj.importTeamspaceProfileTranslateType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE teamspace_profilestranslate");
                        SessionObj.importExcelResult.Add("[TeamspaceProfilesTranslate] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("profileID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM teasmspace_profilestranslate WHERE profileID=@profileID", parameterList);
                        }

                        TProfiles.TProfileSetTranslate tempRight = new TProfiles.TProfileSetTranslate();
                        string tempProfileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        tempRight.profileID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1));
                        tempRight.AllowTranslateRead = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                        tempRight.AllowTranslateEdit = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 3)));
                        tempRight.AllowTranslateRelease = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 4)));
                        tempRight.AllowTranslateNew = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 5)));
                        tempRight.AllowTranslateDelete = getBool(Convert.ToString(tempxls.GetCellValue(rowIndex, 6)));
                        if (!TProfiles.saveProfileTeamspaceTranslate(tempRight, SessionObj.Project.ProjectID))
                            SessionObj.importExcelResult.Add("[TeamspaceProfilesTranslate: " + tempProfileID + "] konnte nicht importiert werden: ProfileID mehrfach vorhanden");

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesTranslate] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesTranslate] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion TeamspaceProfilesTranslate

        # region TeamspaceLicences

        //### Lizenzen können nur geupdatet werden, nicht ersetzt


        sheetIndex = tempxls.GetSheetIndex("Licences", false);
        if (sheetIndex > 0)
        {

            if (SessionObj.importLicences)
            {
                SessionObj.ProcessMessage = "Licences wird importiert: ";
                TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, 0, 0, 0, SessionObj.Project.ProjectID);
                // neue Werte eintragen
                tempxls.ActiveSheet = sheetIndex;
                if ((string)tempxls.GetCellValue(1, 1) == "LicenceID")
                {
                    int selectedIndex = SessionObj.importLicencesType;
                    if (selectedIndex == 0)
                    {
                        // bisherige Werte löschen
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQL("TRUNCATE TABLE licences");
                        SessionObj.importExcelResult.Add("[Licences] vorhandene Werte gelöscht");
                    }

                    TParameterList parameterList = new TParameterList();
                    dataReader = new SqlDB(SessionObj.Project.ProjectID);
                    int rowIndex = 2;
                    while (Convert.ToString(tempxls.GetCellValue(rowIndex, 1)) != "")
                    {
                        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 2, rowIndex - 1, 0, 0, SessionObj.Project.ProjectID);
                        if (selectedIndex == 2)
                        {
                            // bisherige Werte löschen
                            parameterList = new TParameterList();
                            parameterList.addParameter("licenceID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim());
                            dataReader = new SqlDB(SessionObj.Project.ProjectID);
                            dataReader.execSQLwithParameter("DELETE FROM licences WHERE licenceID=@licenceID", parameterList);
                        }

                        string templLicenceID = Convert.ToString(tempxls.GetCellValue(rowIndex, 1)).Trim();
                        // speichern
                        parameterList = new TParameterList();
                        parameterList.addParameter("licenceID", "string", Convert.ToString(tempxls.GetCellValue(rowIndex, 1)));
                        parameterList.addParameter("value", "int", Convert.ToString(tempxls.GetCellValue(rowIndex, 2)));
                        dataReader = new SqlDB(SessionObj.Project.ProjectID);
                        dataReader.execSQLwithParameter("INSERT INTO licences (ID, value) VALUES (@licenceID, @value)", parameterList);

                        rowIndex++;
                    }
                    SessionObj.importExcelResult.Add("[Licences] Import abgeschlossen");
                }
                else
                {
                    // Format der Tabelle fehlerhaft
                    SessionObj.importExcelResult.Add("[TeamspaceProfilesTranslate] konnte nicht importiert werden: Tabellenformat fehlerhaft");
                }
            }
        }
        # endregion TeamspaceLicences

        // Excel-Datei löschen
        File.Delete(SessionObj.importFile);

        // Prozessende eintragen
        TProcessStatus.updateProcessInfo(SessionObj.ProcessID, 0, 0, 0, 0, SessionObj.Project.ProjectID);
    }
    # endregion User-Import
}