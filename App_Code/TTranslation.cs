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
public class TTranslation
{
    public class TApp
    {
        public string name;
        public ArrayList subfolder;
    }
    public class TSubfolder
    {
        public string name;
        public ArrayList objects;
    }
    public class TObject
    {
        public int ID;
        public string filename;
        public string action;
    }
    public ArrayList appList;

    public TTranslation(string aPath, string aProjectID)
    {
        // Struktur aller möglichen Dateien einlesen
        appList = new ArrayList();
        System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
        doc = new System.Xml.XmlDocument();
        doc.Load(aPath);
        // Schleife über alle Folder der XML-Datei
        int actID = 1;
        IEnumerator appIndex = doc.DocumentElement.GetEnumerator();
        System.Xml.XmlElement tempApp;
        while (appIndex.MoveNext())
        {
            tempApp = (System.Xml.XmlElement)appIndex.Current;

            TApp actApp = new TApp();
            actApp.name = tempApp.Attributes.GetNamedItem("name").Value;
            actApp.subfolder = new ArrayList();
            appList.Add(actApp);

            // Schleife über alle Subfolder in der XML-Datei
            System.Xml.XmlElement tempSubfolder;
            IEnumerator folderIndex;
            folderIndex = tempApp.GetElementsByTagName("subfolders")[0].GetEnumerator();
            while (folderIndex.MoveNext())
            {
                tempSubfolder = (System.Xml.XmlElement)folderIndex.Current;

                TSubfolder actSubfolder = new TSubfolder();
                actSubfolder.name = tempSubfolder.Attributes.GetNamedItem("name").Value;
                actSubfolder.objects = new ArrayList();
                actApp.subfolder.Add(actSubfolder);

                // Schleife über alle Subfolder in der XML-Datei
                System.Xml.XmlElement tempObject;
                IEnumerator objectIndex;
                objectIndex = tempSubfolder.GetElementsByTagName("objects")[0].GetEnumerator();
                while (objectIndex.MoveNext())
                {
                    tempObject = (System.Xml.XmlElement)objectIndex.Current;

                    TObject actObject = new TObject();
                    actObject.filename = tempObject.Attributes.GetNamedItem("filename").Value;
                    actObject.action = tempObject.Attributes.GetNamedItem("action").Value;
                    actObject.ID = actID;
                    actSubfolder.objects.Add(actObject);

                    actID++;
                }
            }
        }
    }
    public ArrayList getObjects(string aApp, string aSubfolder)
    {
        ArrayList result = new ArrayList();

        foreach (TApp tempApp in appList)
        {
            if (tempApp.name == aApp)
                foreach (TSubfolder tempSubfolder in tempApp.subfolder)
                {
                    if (tempSubfolder.name == aSubfolder)
                    {
                        foreach (TObject tempObject in tempSubfolder.objects)
                        {
                            result.Add(tempObject.filename);
                        }
                    }
                }
        }

        return result;
    }
    public string getAction(string aToolname, string aCategorie, string aFilename)
    {
        string result = "";

        foreach (TApp tempApp in appList)
        {
            if (tempApp.name == aToolname)
                foreach (TSubfolder tempSubfolder in tempApp.subfolder)
                {
                    if (tempSubfolder.name == aCategorie)
                    {
                        foreach (TObject tempObject in tempSubfolder.objects)
                        {
                            if (tempObject.filename == aFilename)
                                result = tempObject.action;
                        }
                    }
                }
        }

        return result;
    }
    public void getElementByID(int aID, out string aToolname, out string aCategorie, out string aFilename)
    {
        aToolname = "";
        aCategorie = "";
        aFilename = "";

        foreach (TApp tempApp in appList)
        {
            foreach (TSubfolder tempSubfolder in tempApp.subfolder)
            {
                foreach (TObject tempObject in tempSubfolder.objects)
                {
                    if (tempObject.ID == aID)
                    {
                        aToolname = tempApp.name;
                        aCategorie = tempSubfolder.name;
                        aFilename = tempObject.filename;
                    }
                }
            }
        }
    }
}