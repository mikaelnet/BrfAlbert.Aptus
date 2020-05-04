using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BrfAlbert.Aptus
{
    public class Exporter
    {
        private const int ColCustomer = 1;
        private const int ColDate = 2;
        private const int ColObject = 3;
        private const int ColUsed = 4;
        private const int ColPrice = 5;
        private const int ColTotal = 6;

        private readonly ExcelPackage _excel;
        private readonly ExcelWorksheet _sheet;
        private int _row;
        private int _dataStartRow = 1;

        private readonly DateTime _month;

        public Exporter(string name, DateTime month)
        {
            _excel = new ExcelPackage();
            _sheet = _excel.Workbook.Worksheets.Add(name);
            _row = 1;

            _month = new DateTime(month.Year, month.Month, 1);
        }

        private void WriteCell(int row, int col, object value, bool bold = false)
        {
            _sheet.Cells[row, col].Value = value;
            _sheet.Cells[row, col].Style.Font.Bold = bold;
        }

        public void WriteBookingHeaders()
        {
            var cell = _sheet.Cells[_row, 1];
            cell.Value = $"Debiteringsunderlag {_month:MMMM yyyy}";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size += 4;

            _row += 2;
            _sheet.Cells[_row, 1].Value = "Bostadsrättsförening:";
            _sheet.Cells[_row, 2].Value = "Brf Albert, 2151";
            _sheet.Cells[_row, 1, _row, 2].Style.Font.Size += 2;
            _sheet.Cells[_row, 2].Style.Font.Bold = true;

            _row += 2;
            _sheet.Column(ColCustomer).Width = 30;
            WriteCell(_row, ColCustomer, "Lägenhet", true);

            WriteCell(_row, ColDate, "Datum", true);
            _sheet.Column(ColDate).Width = 20;
            _sheet.Column(ColDate).Style.Numberformat.Format = "yyyy-MM-dd (ddd)";
            _sheet.Column(ColDate).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteCell(_row, ColObject, "Lokal", true);
            _sheet.Column(ColObject).Width = 15;

            WriteCell(_row, ColUsed, "Använd", true);
            _sheet.Cells[_row, ColUsed].AddComment("Anger om medlemmen gått in i lokalen med egen tagg. Lokalen kan även ha används om passage t ex skett med nyckel. Bokningen debiteras oavsett.", "Aptus");
            _sheet.Column(ColUsed).Width = 7;

            WriteCell(_row, ColPrice, "Pris", true);
            _sheet.Cells[_row, ColPrice].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            _sheet.Column(ColPrice).Width = 10;
            _sheet.Column(ColPrice).Style.Numberformat.Format = "#,##0 kr";

            WriteCell(_row, ColTotal, "Total", true);
            _sheet.Cells[_row, ColTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            _sheet.Column(ColTotal).Width = 10;
            _sheet.Column(ColTotal).Style.Numberformat.Format = "#,##0 kr";
            _sheet.Column(ColTotal).Style.Font.Bold = true;
            _sheet.Row(_row).Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            _row++;
            _sheet.View.FreezePanes(_row, 1);
            _dataStartRow = _row;
        }

        public void WriteBookingRecords(IEnumerable<CustomerBookings> customers)
        {
            foreach (var customer in customers.Where(c => c.ChargeableBookings.Any()))
            {
                WriteCell(_row, ColCustomer, customer.CustomerName);
                //WriteCell(_row, ColTotal, customer.Total);
                _sheet.Cells[_row, ColTotal].Formula = $"=SUM({_sheet.Cells[_row, ColPrice].Address}:{_sheet.Cells[_row + customer.ChargeableBookings.Count()-1, ColPrice].Address})";
                foreach (var booking in customer.ChargeableBookings)
                {
                    WriteCell(_row, ColDate, booking.PassDate);
                    WriteCell(_row, ColObject, booking.ObjectName);
                    WriteCell(_row, ColUsed, booking.Used ? "Ja" : "Nej");
                    WriteCell(_row, ColPrice, booking.Price);
                    _row++;
                }
                _row++;
            }
        }

        public void WriteBookingFooter(IEnumerable<CustomerBookings> customers)
        {
            WriteCell(_row, 1, $"Skapad: {DateTime.Now:yyyy-MM-dd HH:mm}");
            WriteCell(_row, ColPrice, "Total:", true);
            _sheet.Row(_row).Style.Border.Top.Style = ExcelBorderStyle.Medium;
            //_sheet.Cells[_row, ColTotal].Value = customers.Sum(c => c.Total);
            _sheet.Cells[_row, ColTotal].Formula = $"=SUM({_sheet.Cells[_dataStartRow, ColTotal].Address}:{_sheet.Cells[_row - 1, ColTotal].Address})";
        }

        public void WriteLaundryHeaders()
        {
            var cell = _sheet.Cells[_row, 1];
            cell.Value = $"Användning tvättstuga {_month:MMMM yyyy}";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size += 4;

            _row += 2;
            _sheet.Column(ColCustomer).Width = 30;
            WriteCell(_row, ColCustomer, "Lägenhet", true);

            WriteCell(_row, ColDate, "Datum", true);
            _sheet.Column(ColDate).Width = 20;
            _sheet.Column(ColDate).Style.Numberformat.Format = "yyyy-MM-dd (ddd)";
            _sheet.Column(ColDate).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteCell(_row, ColObject, "Lokal", true);
            _sheet.Column(ColObject).Width = 15;

            WriteCell(_row, ColUsed, "Använd", true);
            _sheet.Cells[_row, ColUsed].AddComment("Anger om medlemmen gått in i lokalen med egen tagg. Lokalen kan även ha används om passage t ex skett med nyckel. Bokningen debiteras oavsett.", "Aptus");
            _sheet.Column(ColUsed).Width = 7;

            _row++;
            _sheet.View.FreezePanes(_row, 1);
        }


        public void WriteLaundryRecords(IEnumerable<CustomerBookings> customers)
        {
            foreach (var customer in customers.Where(c => c.Bookings.Any(b => b.ObjectId == BookingRecord.Tvattstuga)))
            {
                WriteCell(_row, ColCustomer, customer.CustomerName);
                foreach (var booking in customer.Bookings.Where(b => b.ObjectId == BookingRecord.Tvattstuga))
                {
                    WriteCell(_row, ColDate, booking.PassDate);
                    WriteCell(_row, ColObject, booking.ObjectName);
                    WriteCell(_row, ColUsed, booking.Used ? "Ja" : "Nej");
                    _row++;
                }
                _row++;
            }
        }

        public void WriteLaundryFooter(IEnumerable<CustomerBookings> customers)
        {
            WriteCell(_row, 1, $"Skapad: {DateTime.Now:yyyy-MM-dd HH:mm}");
            WriteCell(_row, ColDate, "Totalt antal bokningar:", true);
            _sheet.Row(_row).Style.Border.Top.Style = ExcelBorderStyle.Medium;

            int count = 0;
            foreach (var customer in customers)
            {
                count += customer.Bookings.Count(b => b.ObjectId == BookingRecord.Tvattstuga);
            }

            _sheet.Cells[_row, ColUsed].Value = count;
        }

        public void Save(Stream stream)
        {
            _excel.SaveAs(stream);
        }
    }
}