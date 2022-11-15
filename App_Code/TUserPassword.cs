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
using System.Collections.Specialized;
using System.IO;
using System.Text;

/// <summary>
/// Klasse zur Verwaltung der Benutzers der Tools
/// </summary>
public class TUserPassword
{
    string projectID;
    public int passwordCounter1;
    public int passwordCounter2;
    public int passwordCounter3;
    public int passwordCounter4;
    public string passwordGroup1;
    public string passwordGroup2;
    public string passwordGroup3;
    public string passwordGroup4;
    public int passwordSystemValid;     // Anzahl der Tage, die ein vom System generiertes Passwort gültig ist
    public int passwordUserValid;       // Anzahl der Tage, die ein Benutzer-Passwort gültig ist
    public int passwordAttempt;         // Anzahl der Fehlversuche, nach denen der Account vorübergehend gesperrt wird
    public int passwordLockTime;        // Anzahl der Minuten der Sperrung, wenn zu viele Fehlversuche aufgetreten sind
    public int passwordHistory;         // Anzahl der letzten Passworte, von denen sich das neue unterscheiden muss
    public int passwordLength;          // Minimale erforderliche Passwortlänge
    public bool passwordComplexity;     // PasswortComplexität gefordert: mindestens ein Klein-oder Großbuchstabe, mindestens eine Ziffer, mindestens ein Sonderzeichen
    public int passwordMinValid;        // Minimales Passwort-Alter bevor es wieder geändert werden lann
    public bool showCaptcha;            // Verwendung eines Captchas bei Anforderung Benutzername/Kennwort

    public TUserPassword(string aProjectID)
    {
        if (aProjectID != null)
        {
            projectID = aProjectID;
            // Code-Definition ermitteln und anzeigen
            SqlDB dataReader = new SqlDB(aProjectID);
            passwordSystemValid = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordSystemValid'").intValue;
            passwordUserValid = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordUserValid'").intValue;
            passwordAttempt = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordAttempt'").intValue;
            passwordLockTime = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordLockTime'").intValue;
            passwordHistory = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwortHistory'").intValue;
            passwordLength = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordLength'").intValue;
            passwordComplexity = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordComplexity'").intValue == 1;
            passwordMinValid = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='passwordMinValid'").intValue;
            showCaptcha = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='showCaptcha'").intValue == 1;
            passwordCounter1 = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='PasswordCounter1'").intValue;
            passwordCounter2 = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='PasswordCounter2'").intValue;
            passwordCounter3 = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='PasswordCounter3'").intValue;
            passwordCounter4 = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='PasswordCounter4'").intValue;
            passwordGroup1 = dataReader.getScalarString("SELECT text FROM config WHERE ID='PasswordTextGroup1'").stringValue;
            passwordGroup2 = dataReader.getScalarString("SELECT text FROM config WHERE ID='PasswordTextGroup2'").stringValue;
            passwordGroup3 = dataReader.getScalarString("SELECT text FROM config WHERE ID='PasswordTextGroup3'").stringValue;
            passwordGroup4 = dataReader.getScalarString("SELECT text FROM config WHERE ID='PasswordTextGroup4'").stringValue;
        }
    }
    public string generatePassword(string aUserID, string aProjectID)
    {
        string result = createPassword();
        while (!validPassword(aUserID, result, aProjectID))
        {
            result = createPassword();
        }
        return result;
    }
    public string createPassword()
    {
        // Passwort erzeugen entsprechend den Richlinien
        Random random = new Random();
        int randomValue = 0;
        int randomPosition = 0;
        string tempPassword = "";
        if (passwordComplexity)
        {
            do
            {
                tempPassword = "";
                // Anzahl Ziffern erzeugen
                for (int i = 0; i < passwordCounter1; i++)
                {
                    randomValue = random.Next(passwordGroup1.Length);
                    randomPosition = random.Next(tempPassword.Length);

                    tempPassword = tempPassword.Insert(randomPosition, passwordGroup1[randomValue].ToString());
                }
                // Anzahl Buchstaben 1 erzeugen und an Zufallsposition einfügen
                for (int i = 0; i < passwordCounter2; i++)
                {
                    randomValue = random.Next(passwordGroup2.Length);
                    randomPosition = random.Next(tempPassword.Length);

                    tempPassword = tempPassword.Insert(randomPosition, passwordGroup2[randomValue].ToString());
                }
                // Anzahl Buchstaben 2 erzeugen und an Zufallsposition einfügen
                for (int i = 0; i < passwordCounter3; i++)
                {
                    randomValue = random.Next(passwordGroup3.Length);
                    randomPosition = random.Next(tempPassword.Length);

                    tempPassword = tempPassword.Insert(randomPosition, passwordGroup3[randomValue].ToString());
                }
                // Anzahl Sonderzeichen erzeugen und an Zufallsposition einfügen
                for (int i = 0; i < passwordCounter4; i++)
                {
                    randomValue = random.Next(passwordGroup4.Length);
                    randomPosition = random.Next(tempPassword.Length);

                    tempPassword = tempPassword.Insert(randomPosition, passwordGroup4[randomValue].ToString());
                }
            }
            while ((tempPassword[0] == '0') || (tempPassword[0] == '+') || (tempPassword[0] == '-'));
            // falls Passwort zu kurz, mit Ziffern auffüllen
            if (tempPassword.Length < passwordLength)
            {
                for (int Index = tempPassword.Length; Index < passwordLength; Index++)
                {
                    randomPosition = random.Next(tempPassword.Length);
                    tempPassword = tempPassword.Insert(randomPosition, random.Next(1, 9).ToString());
                }
            }
        }
        else
        {
            for (int Index = 0; Index < passwordLength; Index++)
            {
                tempPassword = tempPassword + random.Next(1, 9).ToString();
            }
        }
        return tempPassword;
    }
    public void updateRules(int aPasswordSystemValid, int aPasswordUserValid, int aPasswordAttempt, int aPasswordLockTime, int aPasswordHistory, int aPasswordLength, bool aPasswordComplexity, int aPasswordMinValid, bool aShowCaptcha, int aGroup1Counter, int aGroup2Counter, int aGroup3Counter, int aGroup4Counter, string aGroup1, string aGroup2, string aGroup3, string aGroup4)
    {
        passwordLength = aPasswordLength;
        passwordMinValid = aPasswordMinValid;
        passwordSystemValid = aPasswordSystemValid;
        passwordUserValid = aPasswordUserValid;
        passwordAttempt = aPasswordAttempt;
        passwordLockTime = aPasswordLockTime;
        passwordHistory = aPasswordHistory;
        passwordComplexity = aPasswordComplexity;
        passwordCounter1 = aGroup1Counter;
        passwordCounter2 = aGroup2Counter;
        passwordCounter3 = aGroup3Counter;
        passwordCounter4 = aGroup4Counter;
        passwordGroup1 = aGroup1;
        passwordGroup2 = aGroup2;
        passwordGroup3 = aGroup3;
        passwordGroup4 = aGroup4;
        showCaptcha = aShowCaptcha;
        string tempComplexity = "0";
        if (passwordComplexity)
            tempComplexity = "1";
        string tempCaptcha = "0";
        if (showCaptcha)
            tempCaptcha = "1";
        setValue(passwordLength.ToString(), "passwordLength", projectID);
        setValue(passwordMinValid.ToString(), "passwordMinValid", projectID);
        setValue(passwordSystemValid.ToString(), "passwordSystemValid", projectID);
        setValue(passwordUserValid.ToString(), "passwordUserValid", projectID);
        setValue(passwordAttempt.ToString(), "passwordAttempt", projectID);
        setValue(passwordLockTime.ToString(), "passwordLockTime", projectID);
        setValue(passwordHistory.ToString(), "passwortHistory", projectID);
        setValue(tempComplexity, "passwordComplexity", projectID);
        setValue(passwordCounter1.ToString(), "PasswordCounter1", projectID);
        setValue(passwordCounter2.ToString(), "PasswordCounter2", projectID);
        setValue(passwordCounter3.ToString(), "PasswordCounter3", projectID);
        setValue(passwordCounter4.ToString(), "PasswordCounter4", projectID);
        setString(passwordGroup1, "PasswordTextGroup1", projectID);
        setString(passwordGroup2, "PasswordTextGroup2", projectID);
        setString(passwordGroup3, "PasswordTextGroup3", projectID);
        setString(passwordGroup4, "PasswordTextGroup4", projectID);
        setValue(tempCaptcha, "showCaptcha", projectID);
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
    // Überprüfung, ob ein Kennwort für enen Benutzer ein zulässig ist
    public bool validPassword(string aUserID, string aPassword, string aProjectID)
    {
        bool result = true;

        // Kennwort auf Mindestlänge testen
        if (aPassword.Length < passwordLength)
            result = false;

        // Kennwort auf Komplexitätsanforderungen testen
        if (passwordComplexity)
        {
            // Gruppe1
            int counter = 0;
            for (int index= 0; index < aPassword.Length; index++)
            {
                if (passwordGroup1.IndexOf(aPassword[index]) >= 0)
                    counter++;
            }
            if (counter < passwordCounter1)
                result = false;

            // Gruppe2
            counter = 0;
            for (int index = 0; index < aPassword.Length; index++)
            {
                if (passwordGroup2.IndexOf(aPassword[index]) >= 0)
                    counter++;
            }
            if (counter < passwordCounter2)
                result = false;

            // Gruppe3
            counter = 0;
            for (int index = 0; index < aPassword.Length; index++)
            {
                if (passwordGroup3.IndexOf(aPassword[index]) >= 0)
                    counter++;
            }
            if (counter < passwordCounter3)
                result = false;

            // Gruppe4
            counter = 0;
            for (int index = 0; index < aPassword.Length; index++)
            {
                if (passwordGroup4.IndexOf(aPassword[index]) >= 0)
                    counter++;
            }
            if (counter < passwordCounter4)
                result = false;
        }

        // Kennwort-Historie prüfen
        bool passwordExits = false;
        int passwordCounter = 0;
        SqlDB dataReader;
        dataReader = new SqlDB("select ID, password from loginpasswordhistory where userID='" + aUserID + "' ORDER BY ID DESC", aProjectID);
        while (dataReader.read() && (passwordCounter < passwordHistory))
        {
            if (TCrypt.validPassword(aPassword, dataReader.getString(1)))
                passwordExits = true;
            passwordCounter++;
        }
        dataReader.close();
        if (passwordExits)
            result = false;

        return result;
    }




    //public string createAndSaveAccessCode(int aEmployeeID)
    //{
    //    string accessCode = createAccessCode();

    //    // Code eintragen
    //    TParameterList parameterList = new TParameterList();
    //    parameterList.addParameter("employeeID", "int", aEmployeeID.ToString());
    //    parameterList.addParameter("accesscode", "string", accessCode);
    //    SqlDB dataReader = new SqlDB(projectID);
    //    dataReader.execSQLwithParameter("UPDATE orgmanager_employee SET accesscode=@accesscode WHERE employeeID=@employeeID", parameterList);

    //    return accessCode;
    //}
}

