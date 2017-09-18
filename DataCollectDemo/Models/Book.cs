using System;
using System.Security.AccessControl;

namespace DataCollectDemo.Models
{
    public class Book
    {
        public Book()
        {
            Ordered = 1;
        }

        //public Int32 Id { get; set; }

        public String Name { get; set; }

        //public String Isbn { get; set; }

        //public String Publisher { get; set; }

        //public String Author { get; set; }

        //public Int32 Price { get; set; }

        public Int32 Ordered { get; set; }

        //public Int32 Received { get; set; }

        //public String Sign { get; set; }

        //public String Note { get; set; }

        public bool Equals(Book other)
        {
            if (other == null) return false;

            if (this.Name == other.Name) return true;
            return false;
        }
    }
}
