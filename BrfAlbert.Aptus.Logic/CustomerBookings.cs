using System;
using System.Collections.Generic;
using System.Linq;

namespace BrfAlbert.Aptus.Logic
{
    public class CustomerBookings
    {

        private readonly List<BookingRecord> _bookings = new List<BookingRecord>();

        public IEnumerable<BookingRecord> Bookings => _bookings;

        public IEnumerable<BookingRecord> ChargeableBookings => _bookings.Where(b => b.Price > 0 && b.ApartmentNumber > 0);

        public string CustomerName => _bookings.FirstOrDefault()?.CustomerName;

        public decimal Total => _bookings.Sum(b => b.Price);

        public void Add(BookingRecord booking)
        {
            _bookings.Add(booking);
        }
    }
}