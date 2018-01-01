using System;
using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class Export
    {
        [Key]
        public int Id { get; set; }
        public string BuildingArea { get; set; }
        public string BuildingName { get; set; }
        public int BuildingNo { get; set; }
        public int Unit { get; set; }
        public int Floor { get; set; }
        public int Room { get; set; }
        public string Name { get; set; }
        public string Call { get; set; }
        public string BrandWidth { get; set; }
        public string ITV { get; set; }
        public string MobilePhone { get; set; }
        public string LinkPhone { get; set; }
        public string TelePhone { get; set; }
        public string Address { get; set; }
    }
}
