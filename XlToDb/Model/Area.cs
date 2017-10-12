using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Area
    {
        public int AreaId { get; set; }

        [StringLength(2)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [StringLength(16)]
        public string Nome { get; set; }
    }
}
