using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using BrfAlbert.Aptus.Logic;

namespace BrfAlbert.Aptus.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            var conStr = ConfigurationManager.ConnectionStrings["Aptus"].ConnectionString;
            var repo = new Repository(conStr);
            var log = System.Console.Out;

            var month = DateTime.Now.AddMonths(-1);
            if (args != null && args.Length > 0)
            {
                if (DateTime.TryParseExact(args[0], "yyyy-MM", CultureInfo.InvariantCulture, 
                    DateTimeStyles.AssumeLocal, out var paramDate))
                    month = paramDate;
            }

            var thisMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            var firstMonth = new DateTime(2019, 10, 1);

            
            log.WriteLine("This application creates an Aptus booking report");
            log.WriteLine("Note: The Aptus MultiAccess application must first be run to let all data be synced");
            log.WriteLine();
            log.WriteLine("Usage:");
            log.WriteLine("  without parameters: Creates a report for the previous month");
            log.WriteLine("  optional month parameter in the form of \"YYYY-MM\": Creates a report for the given month");
            log.WriteLine();

            if (!repo.IsAptusDatabaseUpToDate())
            {
                log.WriteLine($"Warning: The events in the database doesn't seem to be up to date. Have MultiAccess synced?");
            }

            if (month >= thisMonth)
            {
                log.WriteLine("Warning: Current or future months may not contain all bookings!");
            }
            else if (month < firstMonth)
            {
                log.WriteLine("Warning: This booking system was not started until 2019-10!");
            }

            var records = repo.GetCustomerBookings(month).ToList();
            log.WriteLine($"Loaded bookings for {records.Count} customers for {month:MMMM yyyy}");

            var exporter = new Exporter($"Brf Albert {month:yyyy-MM}", month);
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

            exporter = new Exporter($@"Brf Albert {month:yyyy-MM}", month);
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

            log.WriteLine("Created report at {0}", outputFile);
        }
    }
}
