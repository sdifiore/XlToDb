using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Serie
    {
        public int Id { get; set; }

        [StringLength(16)]
        [Display(Name = "Série")]
        public string Apelido { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
