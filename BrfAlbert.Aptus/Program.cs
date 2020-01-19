using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;

namespace BrfAlbert.Aptus
{
    class Program
    {
        static void Main(string[] args)
        {
            var conStr = ConfigurationManager.ConnectionStrings["Aptus"].ConnectionString;
            var repo = new Repository(conStr);

            var month = DateTime.Now.AddMonths(-1);
            if (args != null && args.Length > 0)
            {
                if (DateTime.TryParseExact(args[0], "yyyy-MM", CultureInfo.InvariantCulture, 
                    DateTimeStyles.AssumeLocal, out var paramDate))
                    month = paramDate;
            }

            var thisMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            var firstMonth = new DateTime(2019, 10, 1);

            Console.WriteLine("This application creates an Aptus booking report");
            Console.WriteLine("Note: The Aptus MultiAccess application must first be run to let all data be synced");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  without parameters: Creates a report for the previous month");
            Console.WriteLine("  optional month parameter in the form of \"YYYY-MM\": Creates a report for the given month");
            Console.WriteLine();

            var lastEvent = repo.GetLatestEventTime();
            if (lastEvent.AddHours(4) < DateTime.Now)
            {
                Console.WriteLine($"Warning: The last event in the database is quite old ({lastEvent:yyyy-MM-dd HH:mm}). Have MultiAccess synced?");
            }
            if (month >= thisMonth)
            {
                Console.WriteLine("Warning: Current or future months may not contain all bookings!");
            }
            else if (month < firstMonth)
            {
                Console.WriteLine("Warning: This booking system was not started until 2019-10!");
            }

            var records = repo.GetCustomerBookings(month).ToList();
            Console.WriteLine($"Loaded bookings for {records.Count} customers for {month:MMMM yyyy}");

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

            Console.WriteLine("Created report at {0}", outputFile);
        }
    }
}
