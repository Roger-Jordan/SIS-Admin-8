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
public class TMultiUser
{
    public class TToolRight
    {
        public string toolID;
        public int available;
    }
    public string UserID;				// eindeutige Kennung des aktuellen Benutzers
    public string DisplayName;	    	// Anzeigename des aktuellen Benutzers
    public string Password;		    	// Kennwort des aktuellen Benutzers
    public DateTime PasswordValidDate;
    public string Email;		    	// Email-Adresse der aktuellen Benutzers
    public string Title;
    public string Name;
    public string Firstname;
    public string Street;
    public string Zipcode;
    public string City;
    public string State;
    public string LastLoginErrorCount;		// Anzahl aktueller nicht erfolgreicher Login des aktuellen Benutzers
    public int ForceChange;			// Benutzer muss bei Anmeldung Kennwort ändern
    public int IsInternalUser;			// interner Benutzer, der keine Lizenz benötigt
    public int Active;					// Aktueller Benutzer ist aktiviert
    public int Language;            // Benutzersprache
    public string Comment;             // Benutzerrolle
    public string GroupID;                 // Benutzergruppe des Benutzers
    // Einheiten-Berechtigungen
    public ArrayList unitRights;
    public string unitRightsSelectedValue;  // 
    public int maxUnitRightID;
    public ArrayList backRights;
    public string backRightsSelectedValue;
    public int maxBackRightID;
    public ArrayList org1Rights;
    public string org1RightsSelectedValue;
    public int maxOrg1RightID;
    public ArrayList back1Rights;
    public string back1RightsSelectedValue;
    public int maxBack1RightID;
    public ArrayList org2Rights;
    public string org2RightsSelectedValue;
    public int maxOrg2RightID;
    public ArrayList back2Rights;
    public string back2RightsSelectedValue;
    public int maxBack2RightID;
    public ArrayList org3Rights;
    public string org3RightsSelectedValue;
    public int maxOrg3RightID;
    public ArrayList back3Rights;
    public string back3RightsSelectedValue;
    public int maxBack3RightID;
    // Tool Berechtigungen
    public ArrayList toolRights;
    // Container-Berechtigungen
    public ArrayList containerRights;
    public string containerRightsSelectedValue;
    public int maxContainerRightID;
    // TaskContactTranslate Berechtigungen
    public TUser.TUserRight userRights;
    // Taskroles Berechtigungen
    public ArrayList taskroleRights;
    // Gruppen Berechtigungen
    public ArrayList usergroupRights;
    // TranslateLanguages Berechtigungen
    public ArrayList translateLanguagesRights;

    public TMultiUser()
    {
        unitRights = new ArrayList();
        backRights = new ArrayList();
        org1Rights = new ArrayList();
        back1Rights = new ArrayList();
        org2Rights = new ArrayList();
        back2Rights = new ArrayList();
        org3Rights = new ArrayList();
        back3Rights = new ArrayList();
        toolRights = new ArrayList();
        containerRights = new ArrayList();
        userRights = new TUser.TUserRight();
        taskroleRights = new ArrayList();
        usergroupRights = new ArrayList();
        translateLanguagesRights = new ArrayList();
    }
    public TUser.TOrgRight getUnitRight(int aID)
    {
        TUser.TOrgRight actUnitRight = null;
        foreach (TUser.TOrgRight tempUnitRight in unitRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TBackRight getBackRight(int aID)
    {
        TUser.TBackRight actUnitRight = null;
        foreach (TUser.TBackRight tempUnitRight in backRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TOrg1Right getUnitRightOrg1(int aID)
    {
        TUser.TOrg1Right actUnitRight = null;
        foreach (TUser.TOrg1Right tempUnitRight in org1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TBack1Right getBack1Right(int aID)
    {
        TUser.TBack1Right actUnitRight = null;
        foreach (TUser.TBack1Right tempUnitRight in back1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TOrg2Right getUnitRightOrg2(int aID)
    {
        TUser.TOrg2Right actUnitRight = null;
        foreach (TUser.TOrg2Right tempUnitRight in org2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TBack2Right getBack2Right(int aID)
    {
        TUser.TBack2Right actUnitRight = null;
        foreach (TUser.TBack2Right tempUnitRight in back2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TOrg3Right getUnitRightOrg3(int aID)
    {
        TUser.TOrg3Right actUnitRight = null;
        foreach (TUser.TOrg3Right tempUnitRight in org3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TBack3Right getBack3Right(int aID)
    {
        TUser.TBack3Right actUnitRight = null;
        foreach (TUser.TBack3Right tempUnitRight in back3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TUser.TContainerRight getContainerRight(int aID)
    {
        TUser.TContainerRight actContainerRight = null;
        foreach (TUser.TContainerRight tempContainerRight in containerRights)
        {
            if (tempContainerRight.ID == aID)
                actContainerRight = tempContainerRight;
        }
        return actContainerRight;
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
    protected string checkUpdate(string aValue, string aField, string aUpdate, string aInvalid, ref TParameterList parameterList)
    {
        parameterList.addParameter(aField, "string", aValue.ToString());

        if (aValue.ToString() == aInvalid)
            return aUpdate;
        else
            if (aUpdate == "")
            return aField + "=@" + aField;
        else
            return aUpdate + ", " + aField + "=@" + aField;
    }
    protected string checkUpdate(DateTime aValue, string aField, string aUpdate, string aInvalid, ref TParameterList parameterList)
    {
        parameterList.addParameter(aField, "datetime", aValue.ToString());

        if (aValue.ToString() == aInvalid)
            return aUpdate;
        else
            if (aUpdate == "")
            return aField + "=@" + aField;
        else
            return aUpdate + ", " + aField + "=@" + aField;
    }
    public int getUnitRightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < unitRights.Count; index++)
        {
            if (((TUser.TOrgRight)unitRights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBackRightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < backRights.Count; index++)
        {
            if (((TUser.TBackRight)backRights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getUnitRightIndexOrg1(int aID)
    {
        int result = -1;
        for (int index = 0; index < org1Rights.Count; index++)
        {
            if (((TUser.TOrg1Right)org1Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBack1RightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < back1Rights.Count; index++)
        {
            if (((TUser.TBack1Right)back1Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getUnitRightIndexOrg2(int aID)
    {
        int result = -1;
        for (int index = 0; index < org2Rights.Count; index++)
        {
            if (((TUser.TOrg2Right)org2Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBack2RightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < back2Rights.Count; index++)
        {
            if (((TUser.TBack2Right)back2Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getUnitRightIndexOrg3(int aID)
    {
        int result = -1;
        for (int index = 0; index < org3Rights.Count; index++)
        {
            if (((TUser.TOrg3Right)org3Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBack3RightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < back3Rights.Count; index++)
        {
            if (((TUser.TBack3Right)back3Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getContainerRightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < containerRights.Count; index++)
        {
            if (((TUser.TContainerRight)containerRights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public void deleteRight(int aID)
    {
        TUser.TOrgRight actUnitRight = null;
        foreach (TUser.TOrgRight tempUnitRight in unitRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            unitRights.Remove(actUnitRight);
    }
    public void deleteBackRight(int aID)
    {
        TUser.TBackRight actUnitRight = null;
        foreach (TUser.TBackRight tempUnitRight in backRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            backRights.Remove(actUnitRight);
    }
    public void deleteOrg1Right(int aID)
    {
        TUser.TOrg1Right actUnitRight = null;
        foreach (TUser.TOrg1Right tempUnitRight in org1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            org1Rights.Remove(actUnitRight);
    }
    public void deleteBack1Right(int aID)
    {
        TUser.TBack1Right actUnitRight = null;
        foreach (TUser.TBack1Right tempUnitRight in back1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            back1Rights.Remove(actUnitRight);
    }
    public void deleteOrg2Right(int aID)
    {
        TUser.TOrg2Right actUnitRight = null;
        foreach (TUser.TOrg2Right tempUnitRight in org2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            org2Rights.Remove(actUnitRight);
    }
    public void deleteBack2Right(int aID)
    {
        TUser.TBack2Right actUnitRight = null;
        foreach (TUser.TBack2Right tempUnitRight in back2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            back2Rights.Remove(actUnitRight);
    }
    public void deleteOrg3Right(int aID)
    {
        TUser.TOrg3Right actUnitRight = null;
        foreach (TUser.TOrg3Right tempUnitRight in org3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            org3Rights.Remove(actUnitRight);
    }
    public void deleteBack3Right(int aID)
    {
        TUser.TBack3Right actUnitRight = null;
        foreach (TUser.TBack3Right tempUnitRight in back3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            back3Rights.Remove(actUnitRight);
    }
    public void deleteContainerRight(int aID)
    {
        TUser.TContainerRight actContainerRight = null;
        foreach (TUser.TContainerRight tempContainerRight in containerRights)
        {
            if (tempContainerRight.ID == aID)
                actContainerRight = tempContainerRight;
        }
        if (actContainerRight != null)
            containerRights.Remove(actContainerRight);
    }
    public void update(string aProjectID, TSessionObj aSessionObj)
    {
        SqlDB dataReader;

        // Update-String konstruieren
        string tempUpdate = "";
        TParameterList parameterList = new TParameterList();
        tempUpdate = checkUpdate(DisplayName, "displayName", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(IsInternalUser, "isinternaluser", tempUpdate, "-1", ref parameterList);
        tempUpdate = checkUpdate(Active, "active", tempUpdate, "-1", ref parameterList);
        tempUpdate = checkUpdate(Password, "password", tempUpdate, ref parameterList, true);
        if (Password != "*")
        {
            //Password wird neu gesetzt, Gültigkeitsdauer muss neu gesetzt werden
            PasswordValidDate = DateTime.Now.AddDays(aSessionObj.Project.userPassword.passwordSystemValid);
            tempUpdate = checkUpdate(PasswordValidDate, "passwordvaliddate", tempUpdate, "", ref parameterList);
        }
        tempUpdate = checkUpdate(Email, "email", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(Title, "title", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(Firstname, "firstname", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(Name, "name", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(Street, "street", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(Zipcode, "zipcode", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(State, "state", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(ForceChange, "forcechange", tempUpdate, "-1", ref parameterList);
        tempUpdate = checkUpdate(LastLoginErrorCount, "lastLoginErrorCount", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(Language, "language", tempUpdate, "-1", ref parameterList);
        tempUpdate = checkUpdate(Comment, "comment", tempUpdate, ref parameterList, false);
        //tempUpdate = checkUpdate(GroupID, "groupID", tempUpdate, ref parameterList, false);
        tempUpdate = checkUpdate(GroupID, "groupID", tempUpdate, "-1", ref parameterList);

        parameterList.addParameter("userID", "string", "");

        // Schleife über alle UserIDs
        int index = 0;
        foreach (string tempUserID in aSessionObj.selectedIDList)
        {
            index++;
            TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

            parameterList.changeParameterValue("userID", tempUserID);
            // Daten schreiben
            if (tempUpdate != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE loginuser SET " + tempUpdate + " WHERE userID=@userID", parameterList);
            }
            // Unit- und Container-Berechtigungen speichern wenn nur ein User bearbeitet wurde
            if (aSessionObj.selectedIDList.Count == 1)
            {
                // bisherige Berechtigugen immer zuvor löschen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserorgid WHERE userID='" + tempUserID + "'");
                foreach (TUser.TOrgRight tempUnitRight in unitRights)
                {
                    TUser.saveUnitRight(tempUserID, tempUnitRight, aProjectID);
                }
                // Back-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserback WHERE userID='" + tempUserID + "'");
                foreach (TUser.TBackRight tempUnitRight in backRights)
                {
                    TUser.saveUnitRightBack(this.UserID, tempUnitRight, aProjectID);
                }
                // Org1-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserorgid1 WHERE userID='" + tempUserID + "'");
                foreach (TUser.TOrg1Right tempUnitRight in org1Rights)
                {
                    TUser.saveUnitRightOrg1(this.UserID, tempUnitRight, aProjectID);
                }
                // Org1-Back-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserback1 WHERE userID='" + tempUserID + "'");
                foreach (TUser.TBack1Right tempUnitRight in back1Rights)
                {
                    TUser.saveUnitRightBack1(this.UserID, tempUnitRight, aProjectID);
                }
                // Org2-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserorgid2 WHERE userID='" + tempUserID + "'");
                foreach (TUser.TOrg2Right tempUnitRight in org2Rights)
                {
                    TUser.saveUnitRightOrg2(this.UserID, tempUnitRight, aProjectID);
                }
                // Org2-Back-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserback2 WHERE userID='" + tempUserID + "'");
                foreach (TUser.TBack2Right tempUnitRight in back2Rights)
                {
                    TUser.saveUnitRightBack2(this.UserID, tempUnitRight, aProjectID);
                }
                // Org3-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserorgid3 WHERE userID='" + tempUserID + "'");
                foreach (TUser.TOrg3Right tempUnitRight in org3Rights)
                {
                    TUser.saveUnitRightOrg3(this.UserID, tempUnitRight, aProjectID);
                }
                // Org3-Back-Berechtigungen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserback3 WHERE userID='" + tempUserID + "'");
                foreach (TUser.TBack3Right tempUnitRight in back3Rights)
                {
                    TUser.saveUnitRightBack3(this.UserID, tempUnitRight, aProjectID);
                }
                // bisherige Berechtigugen löschen
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapusercontainer WHERE userID='" + tempUserID + "'");
                foreach (TUser.TContainerRight tempContainerRight in containerRights)
                {
                    TUser.saveContainerRight(tempUserID, tempContainerRight, aProjectID);
                }
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapuserrights WHERE userID='" + tempUserID + "'");
                TUser.saveUserRight(tempUserID, userRights, aProjectID);

                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapusertaskrole WHERE userID='" + tempUserID + "'");
                foreach (int tempTaskRoleRight in taskroleRights)
                {
                    TUser.saveTaskRoleRight(tempUserID, tempTaskRoleRight, aProjectID);
                }

                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapusergroup WHERE userID='" + tempUserID + "'");
                foreach (string tempUsergroupRight in usergroupRights)
                {
                    TUser.saveGroupRight(tempUserID, tempUsergroupRight, aProjectID);
                }

                dataReader = new SqlDB(aProjectID);
                dataReader.execSQL("DELETE FROM mapusertranslatelanguage WHERE userID='" + tempUserID + "'");
                foreach (int tempTranslateLanguageRight in translateLanguagesRights)
                {
                    TUser.saveTranslateLanguageRight(tempUserID, Convert.ToInt32(tempTranslateLanguageRight), aProjectID);
                }
            }
            // Tool Berechtigungen
            foreach (TToolRight tempTool in toolRights)
            {
                TParameterList parameterListTool = new TParameterList();
                parameterListTool.addParameter("userID", "string", tempUserID);
                parameterListTool.addParameter("toolID", "string", tempTool.toolID);
                // Available
                if ((tempTool.available == 0) | (tempTool.available == 1))
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("DELETE FROM mapusertool WHERE userID=@userID AND toolID=@toolID", parameterListTool);
                }

                if (tempTool.available == 1)
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("INSERT INTO mapusertool (userID, toolID) VALUES (@userID, @toolID)", parameterListTool);
                }
            }
        }
    }
    public bool hasContainerRights(int aContainerID)
    {
        bool result = false;
        foreach (TUser.TContainerRight tempContainerRight in containerRights)
        {
            if (tempContainerRight.ContainerID == aContainerID)
                result = true;
        }
        return result;
    }
}

