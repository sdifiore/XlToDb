using System.ComponentModel.DataAnnotations;

namespace XlToDb
{
    public class GrupoRateio
    {
        public int GrupoRateioId { get; set; }

        [StringLength(50)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
