using System.Web.UI;


namespace SIS.Helper
{
    public static class ControlCollectionExtension
    {
        public static void RemoveLast(this ControlCollection controlCollection)
        {
            if (controlCollection.Count < 1)
                return;
            controlCollection.RemoveAt(controlCollection.Count - 1);
        }
    }
}