using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Setor
    {
        public int SetorId { get; set; }

        [StringLength(3)]
        [Display(Name = "Setor")]
        public string Codigo { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
