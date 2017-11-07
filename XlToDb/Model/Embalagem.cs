using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Embalagem
    {
        public int Id { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
