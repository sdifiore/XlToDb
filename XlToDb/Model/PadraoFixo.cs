using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PadraoFixo
    {
        public int Id { get; set; }

        [Display(Name = "Descrição")]
        [StringLength(128)]
        public string Descricao { get; set; }

        public float Valor { get; set; }
    }
}
