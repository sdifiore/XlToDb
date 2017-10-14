using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Pcp
    {
        public int PcpId { get; set; }

        [StringLength(2)]
        [Display(Name = "PCP")]
        public string Descricao { get; set; }
    }
}
