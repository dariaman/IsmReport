using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IsmReport.Models
{
    [Table("invoice")]
    public class InvoiceModel
    {
        [Key]
        public int id { get; set; }
        public string InvoiceNo { get; set; }
        public DateTime? InvoiceDate { get; set; }
        public string PeriodeBln { get; set; }
        public string PeriodeThn { get; set; }
        public string Deskripsi { get; set; }
        public int Qty { get; set; }
        public Decimal GrandTotal { get; set; }
        public string Status { get; set; }
        public string Filename { get; set; }        
        public DateTime CreateDate { get; set; }
        public DateTime? UpdateDate { get; set; }
    }
}