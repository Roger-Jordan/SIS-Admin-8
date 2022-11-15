using System.Text;


namespace SIS.Helper
{
    public abstract class ControlHelper
    {
        // Ersetzt jedes Zeichen in name, welches nicht in [a-zA-Z0-9_] steckt durch einen
        // Underscore. So wird sicherg
        public static string EncodeName(string name)
        {
            StringBuilder result = new StringBuilder(name);
            for (int i = 0; i < result.Length; i++)
            {
                char c = result[i];
                if ((c < 'a' || c > 'z') && (c < 'A' || c > 'Z') && (c != '_') && (c < '0' || c > '9'))
                {
                    result[i] = '_';
                }
            }

            return result.ToString();
        }
    }
}