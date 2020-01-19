using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace BrfAlbert.Aptus
{
    public class BookingRecord
    {
        public const int Foreningslokal = 6;
        public const int Gastrum1 = 7;
        public const int Gastrum2 = 8;
        public const int Tvattstuga = 10;
        public const int SpaBastu = 11;

        public int BookingId { get; }
        public DateTime PassDate { get; }
        public bool Used { get; }
        public bool Released { get; }
        public int ObjectId { get; }
        public string ObjectName { get; }
        public int CustomerId { get; }
        public string CustomerName { get; }

        private int? _apartmentNumber;
        public int ApartmentNumber
        {
            get
            {
                if (_apartmentNumber != null)
                    return (int)_apartmentNumber;
                var re = new Regex("Lgh ([0-9]+)($|[^0-9])");
                var matches = re.Match(CustomerName);
                _apartmentNumber = matches.Success ? _apartmentNumber = int.Parse(matches.Groups[1].Value) : 0;
                return (int)_apartmentNumber;
            }
        }

        public int PriceGroupId { get; set; }
        public decimal Price { get; set; }

        public BookingRecord(int bookingId, DateTime passDate, bool used, bool released, int objectId, string objectName, int customerId, string customerName, int priceGroupId, decimal price)
        {
            BookingId = bookingId;
            PassDate = passDate;
            Used = used;
            Released = released;
            ObjectId = objectId;
            ObjectName = objectName;
            CustomerId = customerId;
            CustomerName = customerName;
            PriceGroupId = priceGroupId;
            Price = price;
        }
    }
}