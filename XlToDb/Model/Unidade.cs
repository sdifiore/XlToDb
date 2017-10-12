using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Unidade
    {
        public int UnidadeId { get; set; }

        [StringLength(2)]
        public string Apelido { get; set; }

        [StringLength(16)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
