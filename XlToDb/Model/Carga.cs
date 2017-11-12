using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Carga
    {
        public int Id { get; set; }

        [StringLength(32)]
        [Display(Name = "Carga")]
        public string Apelido { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
