using System.ComponentModel.DataAnnotations;
using XlToDb.Models;

namespace XlToDb.Model
{
    public class PlanejProducao
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        [Display(Name = "Código")]
        public Produto Produto { get; set; }

        public int PlanejVendaId { get; set; }

        public PlanejVenda PlanejVenda { get; set; }

        [Display(Name = "Produção Mensal -11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos11 { get; set; }

        [Display(Name = "Produção Mensal -10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos10 { get; set; }

        [Display(Name = "Produção Mensal -9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos09 { get; set; }

        [Display(Name = "Produção Mensal -8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos08 { get; set; }

        [Display(Name = "Produção Mensal -7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos07 { get; set; }

        [Display(Name = "Produção Mensal -6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos06 { get; set; }

        [Display(Name = "Produção Mensal -5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos05 { get; set; }

        [Display(Name = "Produção Mensal -4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos04 { get; set; }

        [Display(Name = "Produção Mensal -3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos03 { get; set; }

        [Display(Name = "Produção Mensal -2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos02 { get; set; }

        [Display(Name = "Produção Mensal -1")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos01 { get; set; }

        [Display(Name = "Produção Mensal -0")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PmpAnoMenos00 { get; set; }

        [Display(Name = "Saldo Final Mês -11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos11 { get; set; }

        [Display(Name = "Saldo Final Mês -10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos10 { get; set; }

        [Display(Name = "Saldo Final Mês -9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos09 { get; set; }

        [Display(Name = "Saldo Final Mês -8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos08 { get; set; }

        [Display(Name = "Saldo Final Mês -7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos07 { get; set; }

        [Display(Name = "Saldo Final Mês -6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos06 { get; set; }

        [Display(Name = "Saldo Final Mês -5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos05 { get; set; }

        [Display(Name = "Saldo Final Mês -4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos04 { get; set; }

        [Display(Name = "Saldo Final Mês -3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos03 { get; set; }

        [Display(Name = "Saldo Final Mês -2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos02 { get; set; }

        [Display(Name = "Saldo Final Mês -1")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos01 { get; set; }

        [Display(Name = "Saldo Final Mês -0")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SfmAnoMenos00 { get; set; }
    }
}
