using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelTest.Models
{
    public class Statistic
    {
        public int StatisticId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime CreationDate { get; set; }
        public string Creator { get; set; }
        public Car Car { get; set; }
    }
}