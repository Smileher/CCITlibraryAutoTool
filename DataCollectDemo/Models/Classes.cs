using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollectDemo.Models
{
    public class Classes
    {
        public Classes()
        {
            //Books = new List<Book>();
            CountOfBooks = new Dictionary<string, int>();
        }

        public String Name { get; set; }

        public String Department { get; set; }

        public String Subjecet { get; set; }

        public String Semester { get; set; }

        public Dictionary<string, int> CountOfBooks { get; set; }

        // public IList<Book> Books { get; set; }
    }
}
