using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Familia
    {
        public int FamiliaId { get; set; }

        [StringLength(3)]
        public string Apelido { get; set; }

        [StringLength(50)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
