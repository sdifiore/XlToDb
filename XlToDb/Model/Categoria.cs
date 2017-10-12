using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Categoria
    {
        public int CategoriaId { get; set; }

        [StringLength(2)]
        public string Apelido { get; set; }

        [StringLength(50)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }
    }
}
