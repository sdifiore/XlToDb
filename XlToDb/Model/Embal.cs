using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Embal
    {
        public int Id { get; set; }

        [StringLength(24)]
        [Display(Name = "Sigla")]
        public string Sigla { get; set; }

        public int InsumoId { get; set; }

        public Insumo Insumo { get; set; }
    }
}
