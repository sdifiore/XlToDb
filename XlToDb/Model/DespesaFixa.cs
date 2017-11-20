using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class DespesaFixa
    {
        public int Id { get; set; }

        [StringLength(128)]
        public string Despesa { get; set; }

        [Display(Name = "Valor Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ValorTotal { get; set; }

        [Display(Name = "Comentário")]
        [StringLength(1024)]
        public string Comentario { get; set; }

        [Display(Name = "Cod Crit")]
        public int CodCrit { get; set; }

        [StringLength(128)]
        [Display(Name = "Critério de Rateio")]
        public string CriterioRateio { get; set; }

        [Display(Name = "Rateio Fitas")]
        public float RateioFitas { get; set; }

        [Display(Name = "Rateio Tubos Cordões Perfis")]
        public float RateioTuboCordaoPerfil { get; set; }

        [Display(Name = "Rateio Fios Gax PTFE Puro")]
        public float RateioFioGaxPtfePuro { get; set; }

        [Display(Name = "Rateio Fios Gax PTFE Graf")]
        public float RateioFioGaxPtfeGraf { get; set; }

        [Display(Name = "Rateio Graxa")]
        public float RateioGraxa { get; set; }

        [Display(Name = "Rateio Sucatas")]
        public float RateioSucatas { get; set; }

        [Display(Name = "Rateio Revenda")]
        public float RateioRevenda { get; set; }

        [Display(Name = "Somas")]
        public float Somas { get; set; }
    }
}
