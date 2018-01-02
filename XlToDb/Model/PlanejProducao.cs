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

        [Display(Name = "Un/UArm")]
        public int UnUArm { get; set; }

        [Display(Name = "Estoque de Reposição")]
        public float EstoqueReposicao { get; set; }

        [StringLength(3)]
        [Display(Name = "Tipo de Produto")]
        public string TipoPcp { get; set; }

        [Display(Name = "Estoque")]
        public float Estoque { get; set; }

        [Display(Name = "Venda mês -11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos11 { get; set; }

        [Display(Name = "Venda mês -10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos10 { get; set; }

        [Display(Name = "Venda mês -9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos09 { get; set; }

        [Display(Name = "Venda mês -8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos08 { get; set; }

        [Display(Name = "Venda mês -7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos07 { get; set; }

        [Display(Name = "Venda mês -6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos06 { get; set; }

        [Display(Name = "Venda mês -5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos05 { get; set; }

        [Display(Name = "Venda mês -4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos04 { get; set; }

        [Display(Name = "Venda mês -3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos03 { get; set; }

        [Display(Name = "Venda mês -2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos02 { get; set; }

        [Display(Name = "Venda mês -1")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos01 { get; set; }

        [Display(Name = "Venda mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VendaMesMenos00 { get; set; }

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
