using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Sequencia
    {
        public int SequenciaId { get; set; }

        [StringLength(2)]
        [Display(Name = "Sequência")]
        public string Descricao { get; set; }
    }
}
