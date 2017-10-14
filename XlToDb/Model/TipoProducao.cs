using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class TipoProducao
    {
        public int TipoProducaoId { get; set; }

        [StringLength(8)]
        [Display(Name = "Tipo de Produção")]
        public string Descricao { get; set; }
    }
}
