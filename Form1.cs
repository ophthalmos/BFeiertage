using System;
using System.ComponentModel;
using System.Text;
using System.Globalization; // CultureInfo
using System.Windows.Forms;
using System.IO;

namespace BFeiertage
{
    public partial class Form1 : Form
    {
        public int startYear, endYear;
        public Form2 f2;
        public Form1()
        {
            InitializeComponent();
            dtpStart.MouseWheel += new MouseEventHandler(Dtp_MouseWheel);
            dtpEnde.MouseWheel += new MouseEventHandler(Dtp_MouseWheel);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { btnListe.Focus(); }
            // Form1.KeyPreview muss auf True gesetzt werden! (default = false)
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        private void Dtp_MouseWheel(object sender, MouseEventArgs e)
        {
            if (!(sender is DateTimePicker dtp)) { return; }
            if (e.Delta > 0 && dtp.Value < dtp.MaxDate)
            {
                dtp.Value = dtp.Value.AddYears(1);
            }
            else if (e.Delta < 0 && dtp.Value > dtp.MinDate)
            {
                dtp.Value = dtp.Value.AddYears(-1);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // DateTimePicker auf Folgejahr einstellen
            dtpEnde.Value = System.DateTime.Now.AddYears(1); // Reihenfolge!
            dtpStart.Value = System.DateTime.Now.AddYears(1); // s. Eingabe()
            // Set the MinDate and MaxDate.
            dtpStart.MinDate = new DateTime(1800, 1, 1);
            dtpEnde.MinDate = new DateTime(1800, 1, 1);
            dtpStart.MaxDate = new DateTime(2999, 12, 31);
            dtpEnde.MaxDate = new DateTime(2999, 12, 31);
            //dtpEnde.Focus();
            //dtpEnde.Select();
        }


        private void GetYears() // wiederholt benutzte Routine
        {
            var startDate = dtpStart.Value;
            startYear = startDate.Year;
            var endDate = dtpEnde.Value;
            endYear = endDate.Year;
        }

        private void AddHolidayItem(DateTime date, string name)
        {
            var item = new ListViewItem(date.ToString("d.M.yyyy"));
            item.SubItems.Add(name);
            listView1.Items.Add(item);
        }

        private void ListHolidays() // ListView mit Daten füllen
        {
            GetYears();
            listView1.Items.Clear();
            for (var i = startYear; i <= endYear; i++)
            {
                try
                {
                    var ostern = Ostern(i); // Ostern nur 1x pro Jahr berechnen

                    if (cbWeiberfast.Checked) { AddHolidayItem(ostern.AddDays(-52), "Weiberfastnacht"); }
                    if (cbRosenMo.Checked) { AddHolidayItem(ostern.AddDays(-48), "Rosenmontag"); }
                    if (cbAscherMi.Checked) { AddHolidayItem(ostern.AddDays(-46), "Aschermittwoch"); }
                    if (cbKarfr.Checked) { AddHolidayItem(ostern.AddDays(-2), "Karfreitag"); }
                    if (cbOsterSo.Checked) { AddHolidayItem(ostern, "Ostern*"); }
                    if (cbChrHim.Checked) { AddHolidayItem(ostern.AddDays(39), "Christi Himmelfahrt"); }
                    if (cbPfingSo.Checked) { AddHolidayItem(ostern.AddDays(49), "Pfingsten*"); }
                    if (cbFroni.Checked) { AddHolidayItem(ostern.AddDays(60), "Fronleichnam"); }
                    if (cbBussBe.Checked || cbAventSo.Checked)  // Buß- und Bettag und Advent (benötigen ggf. eigene Berechnung)
                    {
                        var busstag = Busstag(i); // Nur berechnen, wenn gebraucht
                        if (cbBussBe.Checked) { AddHolidayItem(busstag, "Buß- und Bettag"); }
                        if (cbAventSo.Checked)
                        {
                            AddHolidayItem(busstag.AddDays(11), "1. Advent");
                            AddHolidayItem(busstag.AddDays(18), "2. Advent");
                            AddHolidayItem(busstag.AddDays(25), "3. Advent");
                            AddHolidayItem(busstag.AddDays(32), "4. Advent");
                        }
                    }
                }
                catch (InvalidOperationException ex) { MessageBox.Show(ex.Message, "Fehler"); }
            }
            btnSpeichern.Enabled = listView1.Items.Count > 0;
            if (btnSpeichern.Enabled) { btnSpeichern.Focus(); }
            else { btnListe.Focus(); } // Fokus zurücksetzen, wenn nichts da ist
        }

        // aus dem Buch "Astronomical Formulae for Calculators" des Belgiers Jean Meeus (1982,
        // Willmann-Bell-Verlag, Richmond, Virginia). Nach den Angaben dort soll sie im Jahr 1876
        // entstanden und in Butcher's "Ecclesiastical Calendar" veröffentlicht worden sein.
        private static DateTime Ostern(int year) // Ostern berechnen
        {
            var c1 = year % 19;
            var c2 = year / 100;
            var c3 = year % 100;
            var c4 = c2 / 4;
            var c5 = c2 % 4;
            var c6 = (c2 + 8) / 25;
            var c7 = (c2 - c6 + 1) / 3;
            var c8 = (19 * c1 + c2 - c4 - c7 + 15) % 30;
            var c9 = c3 / 4;
            var c10 = c3 % 4;
            var c11 = (32 + 2 * c5 + 2 * c9 - c8 - c10) % 7;
            var c12 = (c1 + 11 * c8 + 22 * c11) / 451;
            var c13 = c8 + c11 - 7 * c12 + 114;
            var c14 = c13 / 31;        // Monat
            var c15 = (c13 % 31) + 1;    // Tag
            var res = new DateTime(year, c14, c15);
            return res;
        }

        private static DateTime Busstag(int jahr) // Buß- & Bettag berechnen
        {
            var wn = new DateTime(jahr, 12, 25); // Weihnacht
            var dw = Convert.ToInt32(wn.DayOfWeek); // Wochentag als Zahl
            dw = (dw == 0) ? 7 : dw; // DayOfWeek 0 = Sunday, 6 = Saturday
            var bt = wn.AddDays(-(dw + 32)); // wn - dw - 4 * 7 - 4
            return bt;
        }

        private void ExpCalendar()
        {
            IFormatProvider culture = new CultureInfo("de-DE", true);
            var yearRange = startYear.ToString(); // // GetYears() wurde bereits in ListHolidays() aufgerufen.
            if (startYear != endYear) { yearRange += "-" + endYear.ToString(); }
            saveFileDialog1.Title = "Datei speichern";
            saveFileDialog1.InitialDirectory = KnownFolders.GetPath(KnownFolder.Downloads);
            saveFileDialog1.FileName = "BFeiertage_" + yearRange + ".ics";
            saveFileDialog1.Filter = "iCalendar (*.ics)|*.ics|Alle Dateien (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var fs = new FileStream(saveFileDialog1.FileName, FileMode.Create, FileAccess.Write))
                    {
                        using (var strmWriter = new StreamWriter(fs, Encoding.UTF8))
                        {
                            strmWriter.WriteLine("BEGIN:VCALENDAR");
                            strmWriter.WriteLine("VERSION:2.0");
                            strmWriter.WriteLine("PRODID:-//BFeiertage.exe//Bewegliche Feiertage " + Application.ProductVersion + "//DE");
                            // strmWriter.WriteLine("PRODID:-//OPHTHALMOSTAR//BFeiertage//DE");
                            strmWriter.WriteLine("METHOD:PUBLISH");
                            for (var i = 0; i < listView1.Items.Count; i++) // alle Zeilen (Items) der ListView
                            {
                                strmWriter.WriteLine("BEGIN:VEVENT");
                                strmWriter.WriteLine("DTSTAMP:" + DateTime.Now.ToString("yyyyMMdd\\THHmmss\\Z"));
                                strmWriter.WriteLine("UID:" + Guid.NewGuid().ToString());
                                for (var j = 0; j < listView1.Items[i].SubItems.Count; j++) // alle Spalten (SubItems) der ListView
                                {
                                    if (j == 0) // 1. Spalte
                                    {
                                        var dtStart = DateTime.ParseExact(listView1.Items[i].SubItems[j].Text, "d.M.yyyy", culture);
                                        strmWriter.WriteLine("DTSTART;VALUE=DATE:" + dtStart.ToString("yyyyMMdd"));
                                        var qs = listView1.Items[i].SubItems[1].Text;
                                        var qi = qs.Contains("*") ? 2 : 1; // mit Oster- oder Pfingstmontag kombinieren
                                        strmWriter.WriteLine("DTEND;VALUE=DATE:" + dtStart.AddDays(qi).ToString("yyyyMMdd"));
                                    }
                                    if (j == 1) // 2. Spalte
                                    {
                                        var inRX = listView1.Items[i].SubItems[j].Text;
                                        char[] xChar = { '*' };
                                        var resRX = inRX.TrimEnd(xChar);
                                        strmWriter.WriteLine("SUMMARY:" + resRX);
                                    }
                                }
                                // 'Transparent' bedeutet: es wird keine Zeit im Kalender blockiert; kein Konlfikt mit anderen Terminen.
                                strmWriter.WriteLine("TRANSP:TRANSPARENT");
                                strmWriter.WriteLine("END:VEVENT");
                            }
                            strmWriter.WriteLine("END:VCALENDAR");
                        }
                    } // strmWriter wird hier geschlossen (schließt fs mit)
                      // fs wird hier geschlossen (falls strmWriter-Erstellung fehlschlug)
                    listView1.Items.Clear();
                    btnSpeichern.Enabled = false;
                    btnListe.Focus();
                }
                catch (Exception ex) when (ex is ArgumentNullException || ex is InvalidOperationException || ex is IOException)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void BtnListe_Click(object sender, EventArgs e)
        {
            ListHolidays(); // Feiertage in ListView schreiben
        }

        private void BtnSpeichern_Click(object sender, EventArgs e)
        {
            ExpCalendar(); // Speichern
        }

        private void Dtp_ValueChanged(object sender, EventArgs e)
        {
            GetYears();
            if (endYear < startYear)
            {
                if (sender == dtpStart) { dtpEnde.Value = dtpStart.Value; }
                else if (sender == dtpEnde) { dtpStart.Value = dtpEnde.Value; }
            }
        }


        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://www.schielen.de");
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Fehler");
            }
        }

        private void ShowHelpForm()
        {
            if (f2 == null || f2.IsDisposed) { f2 = new Form2(); }
            if (f2.WindowState == FormWindowState.Minimized) { f2.WindowState = FormWindowState.Normal; }
            f2.Show(); // Zeigt das Formular
            f2.Activate();
        }

        private void Form1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            hlpevent.Handled = true;
            ShowHelpForm();
        }

        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            ShowHelpForm();
        }
    }
}
