using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;


public class Assets
{
    public enum JavaScriptFile
    {
        DateTimePicker,
        DetailsSidebar,
        ExtTable,
        PureFileUpload,
        TreeListView,
        SearchBox,

        AccountManager,
        Download,
        FollowUp,
        OrgManager,
        Response,
        TeamSpace,
    }

    private static string JavaScriptFilesToFilename(JavaScriptFile file)
    {
        Dictionary<JavaScriptFile, string> files = new Dictionary<JavaScriptFile, string>() {
            { JavaScriptFile.DateTimePicker,  "components/jquery.simple-dtpicker.min.js" },
            { JavaScriptFile.DetailsSidebar,  "components/details-sidebar.min.js" },
            { JavaScriptFile.ExtTable,        "components/ext-table.min.js" },
            { JavaScriptFile.PureFileUpload,  "components/pure-file-upload.min.js" },
            { JavaScriptFile.SearchBox,       "components/search-box.min.js" },
            { JavaScriptFile.TreeListView,    "components/tree-list-view.min.js" },

            { JavaScriptFile.AccountManager,  "tools/accountmanager.min.js" },
            { JavaScriptFile.Download,        "tools/download.min.js" },
            { JavaScriptFile.FollowUp,        "tools/followup.min.js" },
            { JavaScriptFile.OrgManager,      "tools/orgmanager.min.js" },
            { JavaScriptFile.Response,        "tools/response.min.js" },
            { JavaScriptFile.TeamSpace,       "tools/teamspace.min.js" },
        };
        return files[file];
    }


    public static void RegisterJavaScriptFile(Page page, JavaScriptFile file)
    {
        RegisterJavaScriptFile(page, JavaScriptFilesToFilename(file));
    }

    public static void RegisterJavaScriptFile(Page page, string filename)
    {
        string script = string.Format("<script src=\"Assets/js/{0}\"></script>", filename);
        ScriptManager.RegisterStartupScript(page, page.GetType(), filename, script, false);
    }

    public static void RegisterJavaScript(Page page, string scriptId, string script)
    {
        ScriptManager.RegisterStartupScript(page, page.GetType(), scriptId, script, true);
    }
}