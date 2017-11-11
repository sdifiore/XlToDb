using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Fabricante
    {
        public int Id { get; set; }

        [StringLength(16)]
        [Display(Name = "Fabricante")]
        public string Apelido { get; set; }

        [StringLength(128)]
        [Display(Name = "Comentário")]
        public string Comentario { get; set; }

    }
}
