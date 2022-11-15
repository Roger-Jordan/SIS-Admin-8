using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using FlexCel.Core;
using FlexCel.XlsAdapter;

public class TTranslateFile
{
    public int ID;				// eindeutige Kennung der Datei
    public int languageID;
    public string toolname;
    public string categorie;
    public string filename;
    public int version;
    public int status;
    public string description;
    public DateTime lastupdate;
    public string action;
    public string userID;
    public string projectID;

    public TTranslateFile(string aProjectID)
    {
        ID = -1;
        languageID = 0;
        toolname = "";
        categorie = "";
        filename = "";
        version = 0;
        status = 0;
        description = "";
        lastupdate = DateTime.Now;
        action = "";
        userID = "";
        projectID = aProjectID;
    }
    public TTranslateFile(int aFileID, string aProjectID)
    {
        SqlDB dataReader;
        dataReader = new SqlDB("select languageID, toolname, categorie, filename, version, status, description, lastupdate, action, userID FROM teamspace_translatefiles WHERE ID='" + aFileID + "'", aProjectID);
        if (dataReader.read())
        {
            ID = aFileID;
            languageID = dataReader.getInt32(0);
            toolname = dataReader.getString(1);
            categorie = dataReader.getString(2);
            filename = dataReader.getString(3);
            version = dataReader.getInt32(4);
            status = dataReader.getInt32(5);
            description = dataReader.getString(6);
            lastupdate = dataReader.getDateTime(7);
            action = dataReader.getString(8);
            userID = dataReader.getString(9);
        }
        else
        {
            ID = -1;
            languageID = 0;
            toolname = "";
            categorie = "";
            filename = "";
            version = 0;
            status = 0;
            description = "";
            lastupdate = DateTime.Now;
            action = "";
            userID = "";
        }
        dataReader.close();
        projectID = aProjectID;
    }
    public bool fileExits()
    {
        bool result = false;
        SqlDB dataReader = new SqlDB("SELECT ID FROM teamspace_translatefiles WHERE languageID='" + languageID.ToString() + "' AND toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND version='0'", projectID);
        if (dataReader.read())
        {
            result = true;
        }
        dataReader.close();

        return result;
    }
    public TTranslateFile getLastReleasedFile()
    {
        TTranslateFile result = null;
        SqlDB dataReader = new SqlDB("SELECT ID FROM teamspace_translatefiles WHERE languageID='" + languageID.ToString() + "' AND toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND version='0' AND status='1'", projectID);
        if (dataReader.read())
        {
            result = new TTranslateFile(dataReader.getInt32(0), projectID);
        }
        dataReader.close();

        return result;
    }
    /// <summary>
    /// Aufnahme einer neuen Datei. Wenn bereits eine Version dieser Datei vorhanden ist, wird diese bereits vorhandene als Version gekennzeichnet.
    /// </summary>
    /// <param name="aUploadedFile">Vollstädiger Pfad und Name der hochgeladenen und temporär gespeicherten Datei; diese wird an den Zielort verschoben.</param>
    public void insert(string aUploadedFile)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        string actLanguage = SessionObj.translateLaguage.getCodeByValue(languageID);
        SqlDB dataReader;
        TParameterList parameterList;

        // Überprüfung, ob bereits eine Version dieser Datei vorhanden ist
        if (fileExits())
        {
            // höchste aktuelle Version ermitteln
            int maxVersion = 0;
            dataReader = new SqlDB("SELECT max(version) FROM teamspace_translatefiles WHERE languageID='" + languageID.ToString() + "' AND toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "'", projectID);
            if (dataReader.read())
            {
                maxVersion = dataReader.getInt32(0);
            }
            dataReader.close();

            // Datei umbenennen in neue höchste Version
            string newFilename = Path.GetFileNameWithoutExtension(filename);
            string extention = Path.GetExtension(filename);
            string path = "C:\\SISFiles\\7.1\\unternehmen7\\teamspace\\translationfiles\\" + toolname + "\\" + categorie + "\\";
            File.Move(path + actLanguage + "_" + filename, path + actLanguage + "_" + newFilename + "_v" + (maxVersion + 1).ToString() + extention);

            // Version in DB aktualisieren
            dataReader = new SqlDB(projectID);
            parameterList = new TParameterList();
            parameterList.addParameter("version", "int", (maxVersion + 1).ToString());
            dataReader.execSQLwithParameter("UPDATE teamspace_translatefiles SET version=@version WHERE languageID='" + languageID.ToString() + "' AND toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND version='0'", parameterList);
        }

        // neue Datei im Filesystem ablegen
        File.Copy(aUploadedFile, "C:\\SISFiles\\7.1\\unternehmen7\\teamspace\\translationfiles\\" + toolname + "\\" + categorie + "\\" + actLanguage + "_" + filename);

        // neue Datei in DB eintragen
        dataReader = new SqlDB(projectID);
        parameterList = new TParameterList();
        parameterList.addParameter("languageID", "int", languageID.ToString());
        parameterList.addParameter("toolname", "string", toolname);
        parameterList.addParameter("categorie", "string", categorie);
        parameterList.addParameter("filename", "string", filename);
        parameterList.addParameter("version", "int", "0");
        parameterList.addParameter("status", "int", "0");
        parameterList.addParameter("description", "string", description);
        parameterList.addParameter("lastupdate", "datetime", DateTime.Now.ToString());
        parameterList.addParameter("action", "string", action);
        parameterList.addParameter("userID", "string", SessionObj.User.UserID);
        ID = dataReader.execSQLwithParameterAndReadID("INSERT INTO teamspace_translatefiles (languageID, toolname, categorie, filename, version, status, description, lastupdate, action, userID) VALUES (@languageID, @toolname, @categorie, @filename, @version, @status, @description, @lastupdate, @action, @userID)", parameterList);
    }
    public TTranslateFile getLastVersion()
    {
        TTranslateFile result = null;

        int lastVersion = 0;
        SqlDB dataReader;
        dataReader = new SqlDB("select max(Version) FROM teamspace_translatefiles WHERE toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND languageID='" + languageID.ToString() + "'", projectID);
        if (dataReader.read())
        {
            lastVersion = dataReader.getInt32(0);
        }
        dataReader.close();

        dataReader = new SqlDB("select ID FROM teamspace_translatefiles WHERE toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND languageID='" + languageID + "' AND version='" + lastVersion.ToString() + "'", projectID);
        if (dataReader.read())
        {
            result = new TTranslateFile(dataReader.getInt32(0), projectID);
        }
        dataReader.close();

        return result;
    }
    public TTranslateFile getActVersion()
    {
        TTranslateFile result = null;

        int lastVersion = 0;
        SqlDB dataReader;
        dataReader = new SqlDB("select max(Version) FROM teamspace_translatefiles WHERE toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND languageID='" + languageID.ToString() + "'", projectID);
        if (dataReader.read())
        {
            lastVersion = dataReader.getInt32(0);
        }
        dataReader.close();

        dataReader = new SqlDB("select ID FROM teamspace_translatefiles WHERE toolname='" + toolname + "' AND categorie='" + categorie + "' AND filename='" + filename + "' AND languageID='" + languageID + "' AND version='0'", projectID);
        if (dataReader.read())
        {
            result = new TTranslateFile(dataReader.getInt32(0), projectID);
        }
        dataReader.close();

        return result;
    }
    public static ArrayList getVersionIDs(string aToolname, string aCategorie, string aFilename, int aLanguageID, string aProjectID)
    {
        ArrayList result = new ArrayList();
        SqlDB dataReader;
        dataReader = new SqlDB("select version FROM teamspace_translatefiles WHERE toolname='" + aToolname + "' AND categorie='" + aCategorie + "' AND filename='" + aFilename + "' AND languageID='" + aLanguageID.ToString() + "' ORDER BY version DESC", aProjectID);
        while (dataReader.read())
        {
            result.Add(dataReader.getInt32(0));
        }
        dataReader.close();
        return result;
    }
    public static string getFilenameWithoutLanguage(string aFilename)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        string result = "";
        int actPosition = aFilename.IndexOf("_");
        result = aFilename.Substring(actPosition+1, aFilename.Length-actPosition-1);

        return result;
    }
    public static int getLangageID(string aFilename)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];
        
        int result = -1;
        int actPosition = aFilename.IndexOf("_");
        string actLanguage = aFilename.Substring(0, actPosition);
        result = SessionObj.translateLaguage.getValueByCode(actLanguage);

        return result;
    }
    public static string getLanguage(string aFilename)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        string result = "";
        int actPosition = aFilename.IndexOf("_");
        result = aFilename.Substring(0, actPosition);

        return result;
    }
    public void release()
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        string languageName = SessionObj.translateLaguage.getCodeByValue(languageID);

        //###
        switch (action)
        {
            case "mailcopy":
                #region mailcopy
                string sourceFile = SessionObj.ProjectDataPath + "\\" + SessionObj.Project.ProjectID + "\\TeamSpace\\TranslationFiles\\" + toolname + "\\" + categorie + "\\" + languageName + "_" + filename;
                string targetFile = SessionObj.ProjectDataPath + "\\" + SessionObj.Project.ProjectID + "\\" + toolname + "\\" + categorie + "\\" + languageName + "_" + filename;
                // Konfigurationsdaten aus Mail-Datei lesen
                if (File.Exists(targetFile))
                {
                    StringCollection fileHeader = new StringCollection();
                    using (StreamReader sr = new StreamReader(targetFile, Encoding.UTF8))
                    {
                        while (!sr.EndOfStream)
                        {
                            fileHeader.Add(sr.ReadLine());
                        }
                    }

                    // Mail-Body Mail lesen
                    if (File.Exists(sourceFile))
                    {
                        StringCollection fileBody = new StringCollection();
                        using (StreamReader sr = new StreamReader(sourceFile, Encoding.UTF8))
                        {
                            while (!sr.EndOfStream)
                            {
                                fileBody.Add(sr.ReadLine());
                            }
                        }

                        // Mail speichern
                        StreamWriter sw = new StreamWriter(targetFile, false, Encoding.UTF8);
                        sw.WriteLine(fileHeader[0]);
                        sw.WriteLine(fileHeader[1]);
                        sw.WriteLine(fileHeader[2]);
                        sw.WriteLine(fileHeader[3]);
                        for (int index = 0; index < fileBody.Count; index++)
                        {
                            sw.WriteLine(fileBody[index]);
                        }
                        sw.Flush();
                        sw.Close();
                    }
                }
                #endregion mailcopy
                break;

            case "pdfandcopy":
                #region pdfcopy
                //###
                
                
                
                
                #endregion pdfcopy
                break;

            case "dictcopy":
                #region dictionary
                //###




                #endregion dictionary
                break;

            case "cleancopy":
                #region cleancopy
                //###




                #endregion cleancopy
                break;

            default:
                break;
        }
    }
    public void setStatus(int aStatus)
    {
        // Status in DB eintragen
        status = 1;
        SqlDB dataReader = new SqlDB(projectID);
        dataReader.execSQL("UPDATE teamspace_translatefiles SET status='" + aStatus.ToString() + "' WHERE ID='" + ID.ToString() + "'");
    }
    public void createEmailTemplate(string aPath, string aLanguage)
    {
        string sourceFile = aPath + "\\" + projectID + "\\" + toolname + "\\" + categorie + "\\" + aLanguage + "_" + filename;
        string targetFile = aPath + "\\" + projectID + "\\TeamSpace\\TempFiles\\" + filename;
        // Daten aus Mail-Datei lesen
        if (File.Exists(sourceFile))
        {
            // Mail-Body Mail lesen
            StringCollection fileBody = new StringCollection();
            using (StreamReader sr = new StreamReader(sourceFile, Encoding.UTF8))
            {
                while (!sr.EndOfStream)
                {
                    fileBody.Add(sr.ReadLine());
                }
            }

            // Mail speichern
            StreamWriter sw = new StreamWriter(targetFile, false, Encoding.UTF8);
            for (int index = 4; index < fileBody.Count; index++)
            {
                sw.WriteLine(fileBody[index]);
            }
            sw.Flush();
            sw.Close();

            // Datei in Übersetzungsbereich übertragen
            insert(targetFile);

            // Datei löschen
            File.Delete(targetFile);

            // Status auf released setzen
            setStatus(1);
        }

    }
    public void createFileTemplate(string aPath, string aLanguage)
    {
        string sourceFile = aPath + "\\" + projectID + "\\" + toolname + "\\" + categorie + "\\" + aLanguage + "_" + filename;
        string targetFile = aPath + "\\" + projectID + "\\TeamSpace\\TempFiles\\" + filename;

        // Daten kopieren

        if (File.Exists(sourceFile))
        {
            // Datei in Zwischendatei kopieren
            File.Copy(sourceFile, targetFile);

            // Datei in Übersetzungsbereich übertragen
            insert(targetFile);

            // Zwischendatei löschen
            File.Delete(targetFile);

            // Status auf released setzen
            setStatus(1);
        }
    }
    public void createDictionaryTemplate(string aPath, string aLanguage, string aMasterLanguage)
    {
        ExcelFile tempxls;
        TFlxFormat fmtBold;
        TFlxFormat fmtUnLocked;
        TSheetProtectionOptions SheetProtectionOptions;

        // Excel-File erzeugen
        string targetFile = aPath + "\\" + projectID + "\\TeamSpace\\TempFiles\\" + filename;

        tempxls = new XlsFile(true);
        tempxls.NewFile(1, TExcelFileFormat.v2010);
        tempxls.ActiveSheet = 1;
        tempxls.SheetName = aLanguage;
        int actRow = 1;

        // Formate für Excel-Zellen definieren
        fmtBold = tempxls.GetCellVisibleFormatDef(1, 1);
        fmtBold.Font.Style = TFlxFontStyles.Bold;
        fmtUnLocked = tempxls.GetCellVisibleFormatDef(1, 1);
        fmtUnLocked.Locked = false;

        // Header erstellen
        tempxls.SetColWidth(1, 2000);
        tempxls.SetCellValue(actRow, 1, "Field");
        tempxls.SetCellFormat(actRow, 1, tempxls.AddFormat(fmtBold));
        tempxls.SetColWidth(2, 20000);
        tempxls.SetCellValue(actRow, 2, "Comment");
        tempxls.SetCellFormat(actRow, 2, tempxls.AddFormat(fmtBold));
        tempxls.SetColWidth(3, 20000);
        tempxls.SetCellValue(actRow, 3, aMasterLanguage);
        tempxls.SetCellFormat(actRow, 3, tempxls.AddFormat(fmtBold));
        tempxls.SetColWidth(4, 20000);
        tempxls.SetCellValue(actRow, 4, aLanguage);
        tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtBold));
        actRow++;

        // DB-Tabelle bestimmen
        string dbTable = "";
        if (toolname == "Start")
            dbTable = "dictionary";
        else
            dbTable = toolname + "_dictionary";

        // Überprüfung, ob bereits eine freigegebene Version vorhanden ist, falls ja, diese als Basis für die Schlüsselwerte verwenden, ansonsten alle Schlüssel verwenden
        TTranslateFile releasedFile = getLastReleasedFile();
        if (releasedFile != null)
        {
            // Excel-Datei öffnen
            string sourceFile = "C:\\SISFiles\\7.1\\unternehmen7\\teamspace\\translationfiles\\" + toolname + "\\" + categorie + "\\" + aLanguage + "_" + filename;
            ExcelFile sourcexls = new XlsFile(sourceFile);

            // Schleife über alle Zeilen
            while ((sourcexls.GetCellValue(actRow, 1) != null) && (sourcexls.GetCellValue(actRow, 1).ToString() != ""))
            {
                // Schlüssel aus freigegebener Datei ermitteln
                string actField = sourcexls.GetCellValue(actRow, 1).ToString();
                // Werte aus DB lesen und in neue Datei schreiben
                tempxls.SetCellValue(actRow, 1, actField);

                SqlDB dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE language='comment' AND field='" + actField + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 2, dataReader1.getString(0));
                dataReader1.close();

                dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE language='" + aMasterLanguage + "' AND field='" + actField + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(0));
                dataReader1.close();

                dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE language='" + aLanguage + "' AND field='" + actField + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 4, dataReader1.getString(0));
                dataReader1.close();
                tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));

                actRow++;
            }
        }
        else
        {
            SqlDB dataReader = new SqlDB("Select field, text FROM " + dbTable + " WHERE language='" + aLanguage + "' ORDER BY position", projectID);
            while (dataReader.read())
            {
                string actField = dataReader.getString(0);

                tempxls.SetCellValue(actRow, 1, actField);
                tempxls.SetCellValue(actRow, 4, dataReader.getString(1));
                tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));

                SqlDB dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE language='comment' AND field='" + actField + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 2, dataReader1.getString(0));
                dataReader1.close();
                // Masterprache
                dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE language='" + aMasterLanguage + "' AND field='" + actField + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(0));
                dataReader1.close();

                actRow++;
            }
            dataReader.close();
        }
        // Blattschutz aktivieren
        SheetProtectionOptions = new TSheetProtectionOptions(false);
        SheetProtectionOptions.Contents = true;
        SheetProtectionOptions.Objects = true;
        SheetProtectionOptions.Scenarios = true;
        SheetProtectionOptions.SelectLockedCells = true;
        SheetProtectionOptions.SelectUnlockedCells = true;
        tempxls.Protection.SetSheetProtection(null, SheetProtectionOptions);

        tempxls.Save(targetFile, TFileFormats.Xlsx);

        // Datei in Übersetzungsbereich übertragen
        insert(targetFile);

        // Zwischendatei löschen
        File.Delete(targetFile);

        // Status auf released setzen
        setStatus(1);
    }
    public void createOrgManagerFieldDictionaryTemplate(string aPath, string aLanguage, string aMasterLanguage)
    {
        ExcelFile tempxls;
        TFlxFormat fmtBold;
        TFlxFormat fmtUnLocked;
        TSheetProtectionOptions SheetProtectionOptions;
        string dbTable = "";

        // Excel-File erzeugen
        string targetFile = aPath + "\\" + projectID + "\\TeamSpace\\TempFiles\\" + filename;

        tempxls = new XlsFile(true);
        tempxls.NewFile(1, TExcelFileFormat.v2010);
        tempxls.ActiveSheet = 1;
        tempxls.SheetName = aLanguage;
        int actRow = 1;

        // Formate für Excel-Zellen definieren
        fmtBold = tempxls.GetCellVisibleFormatDef(1, 1);
        fmtBold.Font.Style = TFlxFontStyles.Bold;
        fmtUnLocked = tempxls.GetCellVisibleFormatDef(1, 1);
        fmtUnLocked.Locked = false;

        // Header erstellen
        tempxls.SetColWidth(1, 2000);
        tempxls.SetCellValue(actRow, 1, "Field");
        tempxls.SetColWidth(2, 20000);
        tempxls.SetCellValue(actRow, 2, "Comment");
        tempxls.SetColWidth(3, 20000);
        tempxls.SetCellValue(actRow, 3, aMasterLanguage);
        tempxls.SetColWidth(4, 20000);
        tempxls.SetCellValue(actRow, 4, aLanguage);
        actRow++;

        // DB-Tabelle bestimmen
        if (action == "dictemployeecopy")
            dbTable = "orgmanager_employeefieldsdictionary";
        if (action == "dictstructurecopy")
            dbTable = "orgmanager_structurefieldsdictionary";
        if (action == "dictstructure1copy")
            dbTable = "orgmanager_structure1fieldsdictionary";
        if (action == "dictstructure2copy")
            dbTable = "orgmanager_structure2fieldsdictionary";

        // Überprüfung, ob bereits eine freigegebene Version vorhanden ist, falls ja, diese als Basis für die Schlüsselwerte verwenden, ansonsten alle Schlüssel verwenden
        TTranslateFile releasedFile = getLastReleasedFile();
        if (releasedFile != null)
        {
            // Excel-Datei öffnen
            string sourceFile = "C:\\SISFiles\\7.1\\unternehmen7\\teamspace\\translationfiles\\" + toolname + "\\" + categorie + "\\" + aLanguage + "_" + filename;
            ExcelFile sourcexls = new XlsFile(sourceFile);

            // Schleife über alle Zeilen
            while ((sourcexls.GetCellValue(actRow, 1) != null) && (sourcexls.GetCellValue(actRow, 1).ToString() != "") && (sourcexls.GetCellValue(actRow, 2) != null) && (sourcexls.GetCellValue(actRow, 2).ToString() != ""))
            {
                // Schlüssel aus freigegebener Datei ermitteln
                string actField = sourcexls.GetCellValue(actRow, 1).ToString();
                string actColumn = sourcexls.GetCellValue(actRow, 2).ToString();
                // Werte aus DB lesen und in neue Datei schreiben
                tempxls.SetCellValue(actRow, 1, actField);
                tempxls.SetCellValue(actRow, 2, actColumn);
                SqlDB dataReader1 = new SqlDB("Select " + actColumn + " FROM " + dbTable + " WHERE fieldID='" + actField + "' AND language='" + aMasterLanguage + "'", projectID);
                if (dataReader1.read())
                {
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(0));
                }
                dataReader1.close();
                dataReader1 = new SqlDB("Select " + actColumn + " FROM " + dbTable + " WHERE fieldID='" + actField + "' AND language='" + aLanguage + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 4, dataReader1.getString(0));
                dataReader1.close();
                tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));

                actRow++;
            }
        }
        else
        {
            SqlDB dataReader = new SqlDB("Select fieldID, title, tooltipp, errorRegEx, errorMandatory, errorRange, errorMaxLength, errorInvalidValue, exportHintHeader, exportHint1, exportHint2, exportHint3 FROM " + dbTable + " WHERE language='" + aLanguage + "' ORDER BY position", projectID);
            while (dataReader.read())
            {
                string actField = dataReader.getString(0);

                SqlDB dataReader1 = new SqlDB("Select fieldID, title, tooltipp, errorRegEx, errorMandatory, errorRange, errorMaxLength, errorInvalidValue, exportHintHeader, exportHint1, exportHint2, exportHint3 FROM " + dbTable + " WHERE language='" + aMasterLanguage + "' ORDER BY position", projectID);
                if (dataReader1.read())
                {
                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "Title");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(1));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(1));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "Tooltipp");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(2));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(2));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "errorRegEx");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(3));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(3));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "errorMandatory");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(4));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(4));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "errorRange");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(5));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(5));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "errorMaxLength");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(6));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(6));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "errorInvalidValue");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(7));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(7));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "exportHintHeader");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(8));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(8));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "exportHint1");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(9));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(9));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "exportHint2");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(10));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(10));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;

                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, "exportHint3");
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(11));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(11));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;
                }
                dataReader1.close();
            }
            dataReader.close();
        }
        // Blattschutz aktivieren
        SheetProtectionOptions = new TSheetProtectionOptions(false);
        SheetProtectionOptions.Contents = true;
        SheetProtectionOptions.Objects = true;
        SheetProtectionOptions.Scenarios = true;
        SheetProtectionOptions.SelectLockedCells = true;
        SheetProtectionOptions.SelectUnlockedCells = true;
        tempxls.Protection.SetSheetProtection(null, SheetProtectionOptions);

        tempxls.Save(targetFile, TFileFormats.Xlsx);

        // Datei in Übersetzungsbereich übertragen
        insert(targetFile);

        // Zwischendatei löschen
        File.Delete(targetFile);

        // Status auf released setzen
        setStatus(1);

    }
    public void createOrgManagerCategoriesDictionaryTemplate(string aPath, string aLanguage, string aMasterLanguage)
    {
        ExcelFile tempxls;
        TFlxFormat fmtBold;
        TFlxFormat fmtUnLocked;
        TSheetProtectionOptions SheetProtectionOptions;
        string dbTable = "";

        // Excel-File erzeugen
        string targetFile = aPath + "\\" + projectID + "\\TeamSpace\\TempFiles\\" + filename;

        tempxls = new XlsFile(true);
        tempxls.NewFile(1, TExcelFileFormat.v2010);
        tempxls.ActiveSheet = 1;
        tempxls.SheetName = aLanguage;
        int actRow = 1;

        // Formate für Excel-Zellen definieren
        fmtBold = tempxls.GetCellVisibleFormatDef(1, 1);
        fmtBold.Font.Style = TFlxFontStyles.Bold;
        fmtUnLocked = tempxls.GetCellVisibleFormatDef(1, 1);
        fmtUnLocked.Locked = false;

        // Header erstellen
        tempxls.SetColWidth(1, 2000);
        tempxls.SetCellValue(actRow, 1, "Field");
        tempxls.SetColWidth(2, 20000);
        tempxls.SetCellValue(actRow, 2, "Comment");
        tempxls.SetColWidth(3, 20000);
        tempxls.SetCellValue(actRow, 3, aMasterLanguage);
        tempxls.SetColWidth(4, 20000);
        tempxls.SetCellValue(actRow, 4, aLanguage);
        actRow++;

        // DB-Tabelle bestimmen
        if (action == "dictemployeecategoriescopy")
            dbTable = "orgmanager_employeefieldscategories";
        if (action == "dictstructurecategoriescopy")
            dbTable = "orgmanager_structurefieldscategories";
        if (action == "dictstructure1categoriescopy")
            dbTable = "orgmanager_structure1fieldscategories";
        if (action == "dictstructure2categoriescopy")
            dbTable = "orgmanager_structure2fieldscategories";

        // Überprüfung, ob bereits eine freigegebene Version vorhanden ist, falls ja, diese als Basis für die Schlüsselwerte verwenden, ansonsten alle Schlüssel verwenden
        TTranslateFile releasedFile = getLastReleasedFile();
        if (releasedFile != null)
        {
            // Excel-Datei öffnen
            string sourceFile = "C:\\SISFiles\\7.1\\unternehmen7\\teamspace\\translationfiles\\" + toolname + "\\" + categorie + "\\" + aLanguage + "_" + filename;
            ExcelFile sourcexls = new XlsFile(sourceFile);

            // Schleife über alle Zeilen
            while ((sourcexls.GetCellValue(actRow, 1) != null) && (sourcexls.GetCellValue(actRow, 1).ToString() != "") && (sourcexls.GetCellValue(actRow, 2) != null) && (sourcexls.GetCellValue(actRow, 2).ToString() != ""))
            {
                // Schlüssel aus freigegebener Datei ermitteln
                string actField = sourcexls.GetCellValue(actRow, 1).ToString();
                string actValue = sourcexls.GetCellValue(actRow, 2).ToString();
                // Werte aus DB lesen und in neue Datei schreiben
                tempxls.SetCellValue(actRow, 1, actField);
                tempxls.SetCellValue(actRow, 2, actValue);
                SqlDB dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE fieldID='" + actField + "' AND value='" + actValue + "' AND language='" + aMasterLanguage + "'", projectID);
                if (dataReader1.read())
                {
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(0));
                }
                dataReader1.close();
                dataReader1 = new SqlDB("Select text FROM " + dbTable + " WHERE fieldID='" + actField + "' AND value='" + actValue + "' AND language='" + aLanguage + "'", projectID);
                if (dataReader1.read())
                    tempxls.SetCellValue(actRow, 4, dataReader1.getString(0));
                dataReader1.close();
                tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));

                actRow++;
            }
        }
        else
        {
            SqlDB dataReader = new SqlDB("Select fieldID, value, text FROM " + dbTable + " WHERE language='" + aLanguage + "' ORDER BY fieldID, value", projectID);
            while (dataReader.read())
            {
                string actField = dataReader.getString(0);
                int actValue = dataReader.getInt32(1);

                SqlDB dataReader1 = new SqlDB("Select fieldID, value, text FROM " + dbTable + " WHERE language='" + aMasterLanguage + "' AND fieldID='" + actField + "' AND value='" + actValue + "' ORDER BY value", projectID);
                if (dataReader1.read())
                {
                    tempxls.SetCellValue(actRow, 1, actField);
                    tempxls.SetCellValue(actRow, 2, actValue);
                    tempxls.SetCellValue(actRow, 3, dataReader1.getString(2));
                    tempxls.SetCellValue(actRow, 4, dataReader.getString(2));
                    tempxls.SetCellFormat(actRow, 4, tempxls.AddFormat(fmtUnLocked));
                    actRow++;
                }
                dataReader1.close();
            }
            dataReader.close();
        }
        // Blattschutz aktivieren
        SheetProtectionOptions = new TSheetProtectionOptions(false);
        SheetProtectionOptions.Contents = true;
        SheetProtectionOptions.Objects = true;
        SheetProtectionOptions.Scenarios = true;
        SheetProtectionOptions.SelectLockedCells = true;
        SheetProtectionOptions.SelectUnlockedCells = true;
        tempxls.Protection.SetSheetProtection(null, SheetProtectionOptions);

        tempxls.Save(targetFile, TFileFormats.Xlsx);

        // Datei in Übersetzungsbereich übertragen
        insert(targetFile);

        // Zwischendatei löschen
        File.Delete(targetFile);

        // Status auf released setzen
        setStatus(1);
    }

    public static void checkForTranslationFile(string aTool, string aCategorie, string aFilename)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TTranslateFile actTranslateFile = new TTranslateFile(SessionObj.Project.ProjectID);
        actTranslateFile.toolname = aTool;
        actTranslateFile.categorie = aCategorie;
        actTranslateFile.filename = TTranslateFile.getFilenameWithoutLanguage(aFilename);
        actTranslateFile.languageID = TTranslateFile.getLangageID(aFilename);
        if (actTranslateFile.fileExits())
        {
            // Basisdaten ermitteln
            actTranslateFile = actTranslateFile.getActVersion();

            //neu hochgeladene als neue releaste Version eintragen
            actTranslateFile.createFileTemplate(SessionObj.ProjectDataPath, TTranslateFile.getLanguage(aFilename));
        }
    }
    public static void checkForTranslationEmail(string aTool, string aCategorie, string aFilename)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TTranslateFile actTranslateFile = new TTranslateFile(SessionObj.Project.ProjectID);
        actTranslateFile.toolname = aTool;
        actTranslateFile.categorie = aCategorie;
        actTranslateFile.filename = TTranslateFile.getFilenameWithoutLanguage(aFilename);
        actTranslateFile.languageID = TTranslateFile.getLangageID(aFilename);
        if (actTranslateFile.fileExits())
        {
            // Basisdaten ermitteln
            actTranslateFile = actTranslateFile.getActVersion();

            //neu hochgeladene als neue releaste Version eintragen
            actTranslateFile.createEmailTemplate(SessionObj.ProjectDataPath, TTranslateFile.getLanguage(aFilename));
        }
    }
    public static void checkForTranslationDictionary(string aTool, string aCategorie, string aFilename)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TTranslateFile actTranslateFile = new TTranslateFile(SessionObj.Project.ProjectID);
        actTranslateFile.toolname = aTool;
        actTranslateFile.categorie = aCategorie;
        actTranslateFile.filename = TTranslateFile.getFilenameWithoutLanguage(aFilename);
        actTranslateFile.languageID = TTranslateFile.getLangageID(aFilename);
        if (actTranslateFile.fileExits())
        {
            // Basisdaten ermitteln
            actTranslateFile = actTranslateFile.getActVersion();

            //neu hochgeladene als neue releaste Version eintragen
            actTranslateFile.createDictionaryTemplate(SessionObj.ProjectDataPath, TTranslateFile.getLanguage(aFilename), SessionObj.Project.masterLanguage);
        }
    }

}
