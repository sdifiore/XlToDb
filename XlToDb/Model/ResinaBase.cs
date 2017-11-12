using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class ResinaBase
    {
        public int Id { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [StringLength(128)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
