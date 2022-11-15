using System;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Web;
using System.ComponentModel;

// https://stackoverflow.com/questions/187482/how-can-i-use-the-button-tag-with-asp-net
namespace SIS.Components
{
    public enum FirstDayOfWeek
    {
        Monday = 1,
        Tuesday = 2,
        Wednesday = 3,
        Thursday = 4,
        Friday = 5,
        Saturday = 6,
        Sunday = 0
    }

    public class DateTimePicker : TextBox
    {
        public const string ValidTimePattern = "^(0[0-9]|1[0-9]|2[0-3]):([0-5][0-9])$";


        private DateTime _initialTimestamp = DateTime.Now; // current
        private bool _autodateOnStart = false;

        // Entspricht C#-Zeitformatsangaben
        // yyyy: Year, 4 digit
        // yy: Year, 2 digit
        // MM: Month, 01-12
        // M: Month, 1-12
        // dd: Day, 01-31
        // d: Day, 1-31
        // hh: Hour using 12-hour clock 01-12
        // h: Hour using 12-hour clock 1-12
        // HH: Hour using 24-hour clock 00-23
        // H: Hour using 24-hour clock 0-23
        // mm: Minute, 00-59
        // m: Minute, 0-59
        // tt: AM/PM
        private string _outputFormat = "dd.MM.yyyy HH:mm";
        private int _minuteInterval = 30; // min: 5, max: 30;
        private FirstDayOfWeek _firstDayOfWeek = FirstDayOfWeek.Monday;
        private bool _todayButton = true;
        private bool _closeButton = true;
        private bool _closeOnSelected = true;
        private bool _dateOnly = false;
        private bool _timeOnly = false;
        private bool _futureOnly = false;
        
        private DateTime? _minDate = null;
        private DateTime? _maxDate = null;
        // Müssen noch in Time umgewandelt werden
        private string _minTime = null; // 00:00
        private string _maxTime = null; // 23:59

        private string _notLowerThanDateTimePickerControlID;

        private string _translationParameter;

        private bool _emptyValueAllowed = false;
        private bool _clientSideReadOnly = false;

        private string _icon = null;



        public DateTime InitialTimestamp    
        {
            get { return this._initialTimestamp; }
            set { this._initialTimestamp = value; }
        }

        public bool AutodateOnStart
        {
            get { return this._autodateOnStart; }
            set { this._autodateOnStart = value; }
        }

        public string OutputFormat
        {
            get { return this._outputFormat; }
            set { this._outputFormat = value; }
        }

        public int MinuteInterval
        {
            get {
                return this._minuteInterval;
            }
            set
            {
                if (value >= 5 && value <= 30)
                {
                    this._minuteInterval = value;
                }
            }
        }

        public FirstDayOfWeek FirstDayOfWeek
        {
            get { return this._firstDayOfWeek; }
            set { this._firstDayOfWeek = value; }
        }

        public bool TodayButton
        {
            get { return this._todayButton; }
            set { this._todayButton = value; }
        }

        public bool CloseButton
        {
            get { return this._closeButton; }
            set { this._closeButton = value; }
        }

        public bool CloseOnSelected
        {
            get { return this._closeOnSelected; }
            set { this._closeOnSelected = value; }
        }

        public bool DateOnly
        {
            get { return this._dateOnly; }
            set { this._dateOnly = value; }
        }

        public bool TimeOnly
        {
            get { return this._timeOnly; }
            set { this._timeOnly = value; }
        }

        public bool FutureOnly
        {
            get { return this._futureOnly; }
            set { this._futureOnly = value; }
        }

        public Nullable<DateTime> MinDate
        {
            get { return this._minDate; }
            set { this._minDate = value; }
        }

        public Nullable<DateTime> MaxDate
        {
            get { return this._maxDate; }
            set { this._maxDate = value; }
        }

        public string MinTime
        {
            get
            {
                return this._minTime;
            }
            set
            {
                Regex regex = new Regex(ValidTimePattern);
                if (regex.IsMatch(value))
                {
                    this._minTime = value;
                }
            }
        }

        public string MaxTime
        {
            get
            {
                return this._maxTime;
            }
            set
            {
                Regex regex = new Regex(ValidTimePattern);
                if (regex.IsMatch(value))
                {
                    this._maxTime = value;
                }
            }
        }

        public string NotLowerThanDateTimePickerControlID
        {
            get { return this._notLowerThanDateTimePickerControlID; }
            set { this._notLowerThanDateTimePickerControlID = value; }
        }

        public bool EmptyValueAllowed
        {
            get { return this._emptyValueAllowed; }
            set { this._emptyValueAllowed = value; }
        }

        public bool ClientSideReadOnly
        {
            get { return this._clientSideReadOnly; }
            set { this._clientSideReadOnly = value; }
        }

        public string Icon
        {
            get { return this._icon; }
            set { this._icon = value; }
        }




        protected override void Render(HtmlTextWriter writer)
        {
            Assets.RegisterJavaScriptFile(this.Page, Assets.JavaScriptFile.DateTimePicker);

            PrepareTranslationParameter();
            GenerateAndRegisterComponentGenerationJavaScript();


            // Wenn das Icon leer ist, dann wird eine Standard-Textbox ausgegeben
            if (string.IsNullOrEmpty(Icon))
            {
                base.Render(writer);
            }
            else
            {
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "input-group");
                writer.RenderBeginTag(HtmlTextWriterTag.Div);

                writer.AddAttribute(HtmlTextWriterAttribute.Class, "input-group-addon si-left no-padding si-" + Icon);
                writer.AddAttribute(HtmlTextWriterAttribute.For, this.ClientID);
                writer.RenderBeginTag(HtmlTextWriterTag.Label);
                writer.RenderEndTag();

                base.Render(writer);

                writer.RenderEndTag();
            }
        }

        private void PrepareTranslationParameter()
        {
            TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];
            if (SessionObj == null)
            {
                this._translationParameter = "";
                return;
            }

            List<string> dayNames = new List<string>
            {
                "TxtDateTimePickerSunday",
                "TxtDateTimePickerMonday",
                "TxtDateTimePickerTuesday",
                "TxtDateTimePickerWednesday",
                "TxtDateTimePickerThursday",
                "TxtDateTimePickerFriday",
                "TxtDateTimePickerSaturday",
            };
            for (int i = 0; i < dayNames.Count; i++)
            {
                dayNames[i] = Labeling.getLabel(dayNames[i], SessionObj.Language, SessionObj.Project.ProjectID);
            }

            List<string> monthNames = new List<string>
            {
                "TxtDateTimePickerJanuary",
                "TxtDateTimePickerFebruary",
                "TxtDateTimePickerMarch",
                "TxtDateTimePickerApril",
                "TxtDateTimePickerMay",
                "TxtDateTimePickerJune",
                "TxtDateTimePickerJuly",
                "TxtDateTimePickerAugust",
                "TxtDateTimePickerSeptember",
                "TxtDateTimePickerOctober",
                "TxtDateTimePickerNovember",
                "TxtDateTimePickerDecember",
            };
            for (int i = 0; i < monthNames.Count; i++)
            {
                monthNames[i] = Labeling.getLabel(monthNames[i], SessionObj.Language, SessionObj.Project.ProjectID);
            }

            _translationParameter =
                  "{"
                + "custom: {"
                + "days: ['" + string.Join("', '", dayNames.ToArray()) + "'],"
                + "months: ['" + string.Join("', '", monthNames.ToArray()) + "'],"
                + "sep: '‒',"
                + "format: '" + ConvertCSharpToJavaScriptDateTimeFormat(this._outputFormat) + "',"
                + "prevMonth: '" + Labeling.getLabel("TxtDateTimePickerPreviousMonth", SessionObj.Language, SessionObj.Project.ProjectID) + "',"
                + "nextMonth: '" + Labeling.getLabel("TxtDateTimePickerNextMonth", SessionObj.Language, SessionObj.Project.ProjectID) + "',"
                + "today: '" + Labeling.getLabel("TxtDateTimePickerToday", SessionObj.Language, SessionObj.Project.ProjectID) + "'"
                + "}"
                + "}";
        }

        private void GenerateAndRegisterComponentGenerationJavaScript()
        {
            // Wenn das Eingabefeld als ReadOnly markiet ist, dann brauchen wir gar keine
            // Funktion auf das Textfeld legen, da es dann nur zur Anzeige da ist.
            if (ReadOnly)
            {
                return;
            }

            string javascript = ";(function($) {";

            // Der Schreibtschutz muss via JavaScript und darf nicht per HTML gesetzt werden,
            // da sonst das Text-Attribut keinen Wert vom Client erhält.
            if (this._clientSideReadOnly)
            {
                javascript += "$('#" + this.ClientID + "').attr('readonly', 'readonly');";
            }
            

            // Verknüpft die Textbox mit der jQuery-Funktionalität
            javascript += "$('#" + this.ClientID + "').appendDtpicker({";
            List<string> parameters = GenerateJavaScriptParameters();
            javascript += string.Join(", ", parameters.ToArray());
            javascript += " });";

            if (!string.IsNullOrEmpty(this._notLowerThanDateTimePickerControlID) && this.Page != null)
            {
                Control targetControl = this.NamingContainer.FindControl(this._notLowerThanDateTimePickerControlID);
                if (targetControl != null)
                {
                    javascript += "$('#" + targetControl.ClientID + "').change(function() {";
                    javascript += "$('#" + this.ClientID + "').handleDtpicker('setMinDate', $('#" + targetControl.ClientID + "').handleDtpicker('getDate'));";
                    javascript += "});";
                }
                
            }
            javascript += "})(jQuery);";

            Assets.RegisterJavaScript(this.Page, "jquery.simple-dtpicker." + this.ClientID, javascript);
        }

        private List<string> GenerateJavaScriptParameters()
        {
            List<string> parameters = new List<string>();

            if (this._autodateOnStart)
            {
                parameters.Add("autodateOnStart: true");
            }
            else
            {
                parameters.Add("autodateOnStart: true");
                DateTime currentValue;
                if (HasValidText())
                {
                    currentValue = ParseTimestamp().Value;
                }
                else
                {
                    currentValue = this._initialTimestamp;
                }

                // Nicht nur die "current"-Option im Javascript muss gesetzt werden, sondern auch der Wert
                // im Eingabefeld, da sonst das Plugin den Wert nicht immer korrekt parst. Zudem ist wichtig
                // dass immer das vollständige Datum ausgegeben wird, nicht nur die Zeit oder das Datum.
                this.Text = currentValue.ToString("yyyy-MM-dd HH:mm");

                parameters.Add(string.Format("current: '{0}'", this.Text));
            }

            parameters.Add(string.Format("minuteInterval: {0}", this._minuteInterval.ToString()));
            parameters.Add(string.Format("firstDayOfWeek: {0}", ((int)this._firstDayOfWeek).ToString()));

            if (!this._todayButton)
            {
                parameters.Add("todayButton: false");
            }
            if (!this._closeButton)
            {
                parameters.Add("closeButton: false");
            }
            parameters.Add(string.Format("closeOnSelected: {0}", this._closeOnSelected.ToString().ToLower()));
            if (this._dateOnly)
            {
                parameters.Add("dateOnly: true");
                parameters.Add(string.Format("dateFormat: '{0}'", ConvertCSharpToJavaScriptDateTimeFormat(this._outputFormat)));
            }
            if (this._timeOnly)
            {
                parameters.Add("timeOnly: true");
                parameters.Add(string.Format("dateFormat: '{0}'", ConvertCSharpToJavaScriptDateTimeFormat(this._outputFormat)));
            }
            if (this._futureOnly)
            {
                parameters.Add("futureOnly: true");
            }

            if (this._minDate.HasValue)
            {
                // Dieser Parameter muss immer im Format yyyy-MM-dd HH:mm angegeben werden, im Gegensatz zu current
                parameters.Add(string.Format("minDate: '{0}'", this._minDate.Value.ToString("yyyy-MM-dd HH:mm")));
            }
            if (this._maxDate.HasValue)
            {
                // Dieser Parameter muss immer im Format yyyy-MM-dd HH:mm angegeben werden, im Gegensatz zu current
                parameters.Add(string.Format("maxDate: '{0}'", this._maxDate.Value.ToString("yyyy-MM-dd HH:mm")));
            }
            if (!string.IsNullOrEmpty(this._minTime))
            {
                parameters.Add(string.Format("minTime: '{0}'", this._minTime));
            }
            if (!string.IsNullOrEmpty(this._maxTime))
            {
                parameters.Add(string.Format("maxTime: '{0}'", this._maxTime));
            }

            parameters.Add("locale: 'custom'");
            parameters.Add(string.Format("externalLocale: {0}", this._translationParameter));

            return parameters;
        }

        public Nullable<DateTime> ParseTimestamp()
        {
            try
            {
                if (string.IsNullOrEmpty(this.Text))
                {
                    return null;
                }
                else
                {
                    return DateTime.ParseExact(this.Text, this._outputFormat, null);
                }                
            }
            catch
            {
                return null;
            }
        }

        public bool HasValidText()
        {
            if (string.IsNullOrEmpty(this.Text))
            {
                return this._emptyValueAllowed;
            }
            else
            {

                try
                {
                    DateTime.ParseExact(this.Text, this._outputFormat, null);
                    return true;
                }
                catch
                {
                    return false;
                }
            }            
        }

        // JS -- DateTimePickerComponent: http://jq-simple-dtpicker-gh-master.herokuapp.com/jquery.simple-dtpicker.html
        // C#: https://docs.microsoft.com/de-de/dotnet/standard/base-types/standard-date-and-time-format-strings?view=netframework-4.7.1
        // YYYY: Year, 4 digit                  =>   yyyy
        // YY: Year, 2 digit                    =>   yy
        // MM: Month, 01-12                     =>   MM             Keine Änderung notwendig
        // M: Month, 1-12                       =>   M              Keine Änderung notwendig
        // DD: Day, 01-31                       =>   dd
        // D: Day, 1-31                         =>   d
        // HH: Hour using 12-hour clock 01-12   =>   hh
        // H: Hour using 12-hour clock 1-12     =>   h
        // hh: Hour using 24-hour clock 00-23   =>   HH
        // h: Hour using 24-hour clock 0-23     =>   H
        // mm: Minute, 00-59                    =>   mm             Keine Änderung notwendig
        // m: Minute, 0-59                      =>   m              Keine Änderung notwendig
        // tt: am/pm                            =>                  Gibt es in C# so nicht; tt kann aber so bleiben, da der Parser wohl nicht case sensitive kennt
        // TT: AM/PM                            =>   tt
        public static string ConvertJavaScriptToCSharpDateTimeFormat(string jsFormat)
        {
            string csFormat = jsFormat;
            
            // Einfache Ersetzungen
            // Es reicht auch aus, die Vier-Zeichen-Angaben auszulassen, da diese automatisch bei den Zwei-Zeichen-Angaben mit ersetzt werden
            // Analog gilt es beim Tages-Format
            csFormat = csFormat
                        .Replace("YY", "yy")
                        .Replace("D", "d")
                        .Replace("TT", "tt");

            // Bei den Stunden-Formaten müssen H => h und h => H gewandelt werden, daher muss das über
            // eine Schleife laufen
            char[] chars = csFormat.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] == 'H')
                {
                    chars[i] = 'h';
                }
                else if (chars[i] == 'h')
                {
                    chars[i] = 'H';
                }
            }

            csFormat = new string(chars);

            return csFormat;
        }

        // C#: https://docs.microsoft.com/de-de/dotnet/standard/base-types/standard-date-and-time-format-strings?view=netframework-4.7.1
        // JS -- DateTimePickerComponent: http://jq-simple-dtpicker-gh-master.herokuapp.com/jquery.simple-dtpicker.html
        // yyyy: Year, 4 digit                  =>   YYYY
        // yy: Year, 2 digit                    =>   YY
        // MM: Month, 01-12                     =>   MM             Keine Änderung notwendig
        // M: Month, 1-12                       =>   M              Keine Änderung notwendig
        // dd: Day, 01-31                       =>   DD
        // d: Day, 1-31                         =>   D
        // hh: Hour using 12-hour clock 01-12   =>   HH
        // h: Hour using 12-hour clock 1-12     =>   H
        // HH: Hour using 24-hour clock 00-23   =>   hh
        // H: Hour using 24-hour clock 0-23     =>   h
        // mm: Minute, 00-59                    =>   mm             Keine Änderung notwendig
        // m: Minute, 0-59                      =>   m              Keine Änderung notwendig
        // tt: AM/PM                            =>   TT
        public static string ConvertCSharpToJavaScriptDateTimeFormat(string csFormat)
        {
            string jsFormat = csFormat;

            // Einfache Ersetzungen
            // Es reicht auch aus, die Vier-Zeichen-Angaben auszulassen, da diese automatisch bei den Zwei-Zeichen-Angaben mit ersetzt werden
            // Analog gilt es beim Tages-Format
            csFormat = csFormat
                        .Replace("yy", "YY")
                        .Replace("d", "D")
                        .Replace("tt", "TT");

            // Bei den Stunden-Formaten müssen h => H und H => h gewandelt werden, daher muss das über
            // eine Schleife laufen
            char[] chars = csFormat.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] == 'h')
                {
                    chars[i] = 'H';
                }
                else if (chars[i] == 'H')
                {
                    chars[i] = 'h';
                }
            }

            csFormat = new string(chars);

            return csFormat;
        }
    }
}