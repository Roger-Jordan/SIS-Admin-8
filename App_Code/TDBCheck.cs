using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Klasse zur Überprüfung von DB-Inhalten
/// </summary>
public class TDBCheck
{
    public class TErrorLanguageEntry
    {
        public int value;
        public string language;
    }
    public class TErrorForeignValueEntry
    {
        public int value;
    }

    /// <summary>
    /// Überprüfung der Vollständigkeit der Values für jede verfügbare Sprache
    /// </summary>
    /// <param name="aErrorList">Rückgabeparameter mit der Fehlerliste</param>
    /// <param name="aSourceTable">zu untersuchende Tabelle</param>
    /// <param name="aSourceCol">zu untersuchende Spalte mit Values</param>
    /// <param name="aLanguageList">Liste der verfügbaren Sprachen</param>
    /// <param name="aProjectID">ProjectID</param>
    public static void checkLanguages(ref ArrayList aErrorList, string aSourceTable, string aSourceCol, TLanguages aLanguageList, string aProjectID)
    {
        TErrorLanguageEntry tempError = null;
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT DISTINCT(" + aSourceCol + ") from " + aSourceTable, aProjectID);
        while (dataReader.read())
        {
            foreach (TLanguages.TEntry actLanguage in aLanguageList.Language)
            {
                int actValue = dataReader.getInt32(0);
                SqlDB dataReader1;
                dataReader1 = new SqlDB("SELECT " + aSourceCol + " from " + aSourceTable + " WHERE value='" + actValue + "' AND language='" + actLanguage.Language + "'", aProjectID);
                if (!dataReader1.read())
                {
                    // Textelement nicht vorhanden
                    tempError = new TErrorLanguageEntry();
                    tempError.value = actValue;
                    tempError.language = actLanguage.Language;
                    aErrorList.Add(tempError);
                }
                dataReader1.close();
            }
        }
        dataReader.close();
    }
    /// <summary>
    /// Überprüfung des Vorhandenseins aller Values als Schlüssel in einer anderen Tabelle
    /// </summary>
    /// <param name="aErrorList">Rückgabeparameter mit der Fehlerliste</param>
    /// <param name="aSourceTable">zu untersuchende Tabelle</param>
    /// <param name="aSourceCol">zu untersuchende Spalte mit Values</param>
    /// <param name="aForeignTable">Tabelle mit allen gültigen Schlüsseln</param>
    /// <param name="aForeignCol">Spalte mit allen gültigen Schlüsseln</param>
    /// <param name="aProjectID">ProjectID</param>
    public static void checkForeignValues(ref ArrayList aErrorList, string aSourceTable, string aSourceCol, string aForeignTable, string aForeignCol, bool nullValid, string aProjectID)
    {
        TErrorForeignValueEntry tempError = null;
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT DISTINCT(" + aSourceCol + ") from " + aSourceTable, aProjectID);
        while (dataReader.read())
        {
            int actValue = dataReader.getInt32(0);
            if (!nullValid || actValue != 0)
            {
                SqlDB dataReader1;
                dataReader1 = new SqlDB("SELECT " + aForeignCol + " from " + aForeignTable + " WHERE " + aForeignCol + "='" + actValue.ToString() + "'", aProjectID);
                if (!dataReader1.read())
                {
                    // Foreign Value nicht vorhanden
                    tempError = new TErrorForeignValueEntry();
                    tempError.value = actValue;
                    aErrorList.Add(tempError);
                }
                dataReader1.close();
            }
        }
        dataReader.close();
    }
}
