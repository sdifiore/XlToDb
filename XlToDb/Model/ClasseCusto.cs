using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class ClasseCusto
    {
        public int ClasseCustoId { get; set; }

        [StringLength(2)]
        [Display(Name = "Classe de Custo")]
        public string Apelido { get; set; }

        [StringLength(50)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
