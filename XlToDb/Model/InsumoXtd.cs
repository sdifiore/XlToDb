using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class InsumoXtd
    {
        public int InsuMoXtdId { get; set; }

        [Display(Name = "Peso")]
        public float Peso { get; set; }

        [Display(Name = "Última atualização")]
        [StringLength(1024)]
        public string Cotacao { get; set; }

        [Display(Name = "Preço USD")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PrecoUsd { get; set; }

        [Display(Name = "Preço Fixado R$")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PrecoRs { get; set; }

        [Display(Name = "ICMS")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float Icms { get; set; }

        [Display(Name = "IPI")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float Ipi { get; set; }

        [Display(Name = "PIS")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float Pis { get; set; }

        [Display(Name = "Cofins")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float Cofins { get; set; }

        [Display(Name = "Despesas Extras")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float DespExtra { get; set; }

        [Display(Name = "II + Desp. Importação")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float DespImport { get; set; }

        [Display(Name = "Ativo")]
        public bool Ativo { get; set; }

        [Display(Name = "Finalidade")]
        public int FinalidadeId { get; set; }

        public Finalidade Finalidade { get; set; }

        [ForeignKey("UnidadeConsumo")]
        [Display(Name = "Unidade de Consumo")]
        public int UnddId { get; set; }

        public Unidade UnidadeConsumo { get; set; }

        [Display(Name = "Quantidade unidades consumo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float QtdUnddConsumo { get; set; }

        [Display(Name = "Quantidade Múltiplo Compra")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float QtdMltplCompra { get; set; }

        [Display(Name = "Forma de Pagamento")]
        [StringLength(16)]
        public string FormaPgto { get; set; }

        [Display(Name = "Prazo 1")]
        public int Prazo1 { get; set; }

        [Display(Name = "Prazo 2")]
        public int Prazo2 { get; set; }

        [Display(Name = "% Pag 1")]
        public float PctPgto1 { get; set; }

        [Display(Name = "Import Pz Pag Desp")]
        public int ImportPzPagDesp { get; set; }

        public int BaseId { get; set; }

        public Base Base { get; set; }
    }
}
