using System;
using System.Collections.Generic;
using System.Text;

    /// <summary>
    /// ein Skalar, d.h. ein einzelner Wert, den eine Datenbankabfrage zurückgibt
    /// </summary>
public class TSkalar
{
    public bool valid = false; //true, wenn die Abfrage einen Wert ergibt
    public int intValue = 0; //wert, wenn ein Integerwert abgefragt wurde
    public string stringValue = ""; //wert, wenn ein String-Wert abgefragt wurde
}