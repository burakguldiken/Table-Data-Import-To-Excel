using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace excelApp
{
    [Table("user")]
    public class User
    {
        public long id { get; set; }
        public string name { get; set; }
        public string surname { get; set; }
        public int age { get; set; }
    }
}
