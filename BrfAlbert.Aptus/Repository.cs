﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace BrfAlbert.Aptus
{
    public class Repository
    {
        private readonly string _connectionString;

        public Repository(string connectionString)
        {
            _connectionString = connectionString;
        }
#if false
SELECT /*bd.id AS BookingId, bd.PassDate, bd.PassNo, bd.Used, bd.Released, 
    bo.Id AS ObjectId, bo.Name AS ObjectName, c.Id AS CustomerId,*/ c.Name AS CustomerName, count(*)
	/*bo.PriceGroup_Id AS PriceGroupId, pgc.Price AS Price*/
FROM BookingData bd
JOIN BookingObject bo ON bd.BookingObject_Id=bo.Id
JOIN Customer c ON bd.CustomerId=c.Id
LEFT OUTER JOIN PriceGroup pg ON bo.PriceGroup_Id=pg.Id
LEFT OUTER JOIN PriceGroupCost pgc ON pg.Id=pgc.PriceGroup_Id
WHERE bd.PassDate between '2019-10-01' and '2019-11-01' /*@from and @until*/
/*AND bd.BookingObject_Id IN (6, 7, 8)*/ AND c.Name LIKE 'Lgh%'
GROUP BY c.Name ORDER BY count(*) DESC
ORDER BY bo.Name, PassDate, PassNo
#endif


        public IEnumerable<BookingRecord> GetBookings(DateTime from, DateTime until)
        {
            const string cmdText = @"
SELECT bd.id AS BookingId, bd.PassDate, bd.PassNo, bd.Used, bd.Released, 
    bo.Id AS ObjectId, bo.Name AS ObjectName, c.Id AS CustomerId, c.Name AS CustomerName,
    COALESCE(bo.PriceGroup_Id, 0) AS PriceGroupId, COALESCE(pgc.Price, 0) AS Price
FROM BookingData bd
JOIN BookingObject bo ON bd.BookingObject_Id=bo.Id
JOIN Customer c ON bd.CustomerId=c.Id
LEFT OUTER JOIN PriceGroup pg ON bo.PriceGroup_Id=pg.Id
LEFT OUTER JOIN PriceGroupCost pgc ON pg.Id=pgc.PriceGroup_Id
WHERE bd.PassDate between @from and @until 
ORDER BY c.Name, PassDate, bo.Name";
            var con = new SqlConnection(_connectionString);
            con.Open();
            var cmd = new SqlCommand(cmdText, con);
            cmd.Parameters.Add(new SqlParameter("@from", SqlDbType.DateTime) {Value = from});
            cmd.Parameters.Add(new SqlParameter("@until", SqlDbType.DateTime) {Value = until});
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    yield return new BookingRecord(
                        reader.GetInt32(reader.GetOrdinal("BookingId")),
                        reader.GetDateTime(reader.GetOrdinal("PassDate")),
                        reader.GetBoolean(reader.GetOrdinal("Used")),
                        reader.GetBoolean(reader.GetOrdinal("Released")),
                        reader.GetInt32(reader.GetOrdinal("ObjectId")),
                        reader.GetString(reader.GetOrdinal("ObjectName")),
                        reader.GetInt32(reader.GetOrdinal("CustomerId")),
                        reader.GetString(reader.GetOrdinal("CustomerName")),
                        reader.GetInt32(reader.GetOrdinal("PriceGroupId")),
                        reader.GetDecimal(reader.GetOrdinal("Price"))
                    );
                }
            }
        }

        public IEnumerable<CustomerBookings> GetCustomerBookings(DateTime month)
        {
            var from = new DateTime(month.Year, month.Month, 1);
            var until = from.AddMonths(1).AddHours(-2);

            CustomerBookings customerBooking = null;
            foreach (var record in GetBookings(from, until).OrderBy(b => b.ApartmentNumber).ThenBy(b => b.PassDate). ThenBy(b => b.ObjectName))
            {
                if (customerBooking == null)
                {
                    customerBooking = new CustomerBookings();
                }
                else if (customerBooking.Bookings.First().CustomerId != record.CustomerId)
                {
                    yield return customerBooking;
                    customerBooking = new CustomerBookings();
                }
                customerBooking.Add(record);
            }

            if (customerBooking != null)
                yield return customerBooking;
        }

        public DateTime GetLatestEventTime()
        {
            const string cmdText = @"SELECT MAX([EventTime]) FROM [Event]";
            var con = new SqlConnection(_connectionString);
            con.Open();
            var cmd = new SqlCommand(cmdText, con);
            return (DateTime) cmd.ExecuteScalar();
        }
    }
}