using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Globalization;

/// <summary>
/// Zusammenfassungsbeschreibung für ConvertDate
/// </summary>
public class TCrypt
{
    // Initialisierungsvektor
    private static readonly byte[] iv = new byte[] { 111, 66, 25, 66, 180, 69, 219, 222 };
    // Key
    private static readonly byte[] key = new byte[] { 9, 12, 10, 16, 25, 9, 4, 25, 17, 1, 13, 21, 19, 5, 19, 13, 21, 6, 2, 7, 8, 14, 13, 18 };

    public TCrypt()
    {
    }

    public static string StringVerschluesseln(string aPassword)
    {
        string salt = BCrypt.Net.BCrypt.GenerateSalt();
        string hash = BCrypt.Net.BCrypt.HashPassword(aPassword + "#Ipsos#", salt);

        return hash;
    }

    public static bool validPassword(string aPassword, string aHash)
    {
        return BCrypt.Net.BCrypt.Verify(aPassword + "#Ipsos#", aHash);
    }
    /// <summary>
    /// Verschlüsselt einen Eingabestring.
    /// </summary>
    /// <param name="input">Der zu verschlüsselnde String.</param>
    /// <returns>Byte-Array mit dem verschlüsselten String.</returns>
    public static string StringVerschluesselnX(string input)
    {
        try
        {
            // MemoryStream Objekt erzeugen
            MemoryStream memoryStream = new MemoryStream();

            // CryptoStream Objekt erzeugen und den Initialisierungs-Vektor
            // sowie den Schlüssel übergeben.
            CryptoStream cryptoStream = new CryptoStream(
            memoryStream, new TripleDESCryptoServiceProvider().CreateEncryptor(key, iv), CryptoStreamMode.Write);

            // Eingabestring in ein Byte-Array konvertieren
            byte[] toEncrypt = new UTF8Encoding().GetBytes(input);

            // Byte-Array in den Stream schreiben und flushen.
            cryptoStream.Write(toEncrypt, 0, toEncrypt.Length);
            cryptoStream.FlushFinalBlock();

            // Ein Byte-Array aus dem Memory-Stream auslesen
            byte[] encrypted = memoryStream.ToArray();

            // Stream schließen.
            cryptoStream.Close();
            memoryStream.Close();

            // Konvertierung in Base64-String
            string encryptedString = Convert.ToBase64String(encrypted);
            
            // Rückgabewert.
            return encryptedString;
        }
        catch (CryptographicException e)
        {
            Console.WriteLine(String.Format(CultureInfo.CurrentCulture, "Fehler beim Verschlüsseln: {0}", e.Message));
            return null;
        }
    }

    /// <summary>
    /// Entschlüsselt einen String aus einem Byte-Array.
    /// </summary>
    /// <param name="data">Das verscghlüsselte Byte-Array.</param>
    /// <returns>Entschlüsselter String.</returns>
    public static string StringEntschluesseln(string data)
    {
        try
        {
            // Verschluesselten Base64String in Byte-Array konvertieren
            byte[] encrypted = Convert.FromBase64String(data);
            
            // Ein MemoryStream Objekt erzeugen und das Byte-Array mit den verschlüsselten Daten zuweisen.
            MemoryStream memoryStream = new MemoryStream(encrypted);

            // Ein CryptoStream Objekt erzeugen und den MemoryStream hinzufügen.
            // Den Schlüssel und Initialisierungsvektor zum entschlüsseln verwenden.
            CryptoStream cryptoStream = new CryptoStream(
            memoryStream,
            new TripleDESCryptoServiceProvider().CreateDecryptor(key, iv), CryptoStreamMode.Read);
            // Buffer erstellen um die entschlüsselten Daten zuzuweisen.
            byte[] fromEncrypt = new byte[data.Length];

            // Read the decrypted data out of the crypto stream
            // and place it into the temporary buffer.
            // Die entschlüsselten Daten aus dem CryptoStream lesen
            // und im temporären Puffer ablegen.
            cryptoStream.Read(fromEncrypt, 0, fromEncrypt.Length);

            // Den Puffer in einen String konvertieren und zurückgeben
            return Encoding.UTF8.GetString(fromEncrypt).Replace("\0", string.Empty);
        }
        catch (CryptographicException e)
        {
            Console.WriteLine(String.Format(CultureInfo.CurrentCulture, "Fehler beim Entschlüsseln: {0}", e.Message));
            return null;
        }
    }
}
