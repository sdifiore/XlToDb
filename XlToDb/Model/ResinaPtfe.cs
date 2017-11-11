using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class ResinaPtfe
    {
        public int Id { get; set; }

        public int Ref { get; set; }

        [StringLength(16)]
        [Display(Name = "Referência")]
        public string Referencia { get; set; }

        [Display(Name = "Fabricante")]
        public int FabricanteId { get; set; }

        public Fabricante Fabricante { get; set; }

        [StringLength(16)]
        public string Tipo { get; set; }

        public int InsumoId { get; set; }

        public Insumo Insumo { get; set; }

        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Custo { get; set; }

        [Display(Name = "Max RR")]
        public int MaxRr { get; set; }

        [StringLength(16)]
        [Display(Name = "Classificação")]
        public string Classificacao { get; set; }

        [Display(Name = "Max RR Antiga")]
        public int MaxRrAntiga { get; set; }
    }
}
