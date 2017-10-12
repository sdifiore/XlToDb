using System.ComponentModel.DataAnnotations;

namespace XlToDb
{
    public class Linha
    {
        public int LinhaId { get; set; }

        [StringLength(4)]
        public string Apelido { get; set; }

        [StringLength(50)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
