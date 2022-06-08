using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BrfAlbert.Aptus.Logic;

namespace BrfAlbert.Aptus.UI
{
    public partial class AptusBillingUI : Form
    {
        private StringWriter _logWriter = new StringWriter();
        private static readonly object _lock = new object();

        public AptusBillingUI()
        {
            InitializeComponent();
            InitializePeriods();
        }

        private void InitializePeriods()
        {
            periodComboBox.Items.Clear();
            var date = DateTime.UtcNow.AddMonths(-1);
            for (int i = 0; i < 10; i++)
            {
                var period = $"{date:yyyy-MM}";
                periodComboBox.Items.Add(period);
                date = date.AddMonths(-1);
            }
            periodComboBox.SelectedIndex = 0;
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void UpdateLogWindow()
        {
            string log;
            lock(_lock)
            {
                log = _logWriter.ToString();
            }

            logTextBox.Invoke((MethodInvoker)delegate { logTextBox.Text = log; });
        }

        public void Log (string message)
        {
            lock(_lock)
            {
                _logWriter.WriteLine(message);
            }

            UpdateLogWindow();
        }

        private void generateButton_Click(object sender, EventArgs e)
        {
            var conStr = ConfigurationManager.ConnectionStrings["Aptus"].ConnectionString;
            if (string.IsNullOrEmpty(conStr))
            {
                Log("Missing database connection string in configuration");
                return;
            }

            var repo = new Repository(conStr);

            if (!repo.IsAptusDatabaseUpToDate())
            {
                Log($"Warning: The events in the database doesn't seem to be up to date. Have MultiAccess synced?");
            }

            var month = GetSelectedMonth();
            if (!month.HasValue)
                return;

            Log($"Loading reservations for {month:MMMM yyyy}");
            var records = repo.GetCustomerBookings(month.Value).ToList();
            Log($"Loaded bookings for {records.Count} customers for {month:MMMM yyyy}");

            var exporter = new Exporter($"Brf Albert {month:yyyy-MM}", month.Value);
            exporter.WriteBookingHeaders();
            exporter.WriteBookingRecords(records);
            exporter.WriteBookingFooter(records);

            var outputFile = $@"BrfAlbert, Debiteringsunderlag {month:yyyy-MM}.xlsx";
            if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["BasePath"]))
            {
                outputFile = Path.Combine(ConfigurationManager.AppSettings["BasePath"], outputFile);
            }
            using (var stream = File.Open(outputFile, FileMode.Create))
            {
                exporter.Save(stream);
            }
            Log($"Created report at {outputFile}");

            exporter = new Exporter($@"Brf Albert {month:yyyy-MM}", month.Value);
            exporter.WriteLaundryHeaders();
            exporter.WriteLaundryRecords(records);
            exporter.WriteLaundryFooter(records);

            outputFile = $@"BrfAlbert, Tvättstuga {month:yyyy-MM}.xlsx";
            if (!string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings["BasePath"]))
            {
                outputFile = Path.Combine(ConfigurationManager.AppSettings["BasePath"], outputFile);
            }
            using (var stream = File.Open(outputFile, FileMode.Create))
            {
                exporter.Save(stream);
            }

            Log($"Created report at {outputFile}");
        }

        private DateTime? GetSelectedMonth()
        {
            if (periodComboBox.SelectedIndex < 0)
            {
                Log("No selected month");
                return null;
            }

            var monthStr = periodComboBox.Items[periodComboBox.SelectedIndex].ToString();
            if (string.IsNullOrEmpty(monthStr))
            {
                Log("Please select a month");
                return null;
            }

            if (!DateTime.TryParseExact(monthStr, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var month))
            {
                Log("Can't parse the selected date");
                return null;
            }

            return month;
        }
    }
}
