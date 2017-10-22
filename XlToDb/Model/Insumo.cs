using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Insumo
    {
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity), Key]
        public int InsumoId { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [Display(Name = "Peso")]
        public float Peso { get; set; }

        [Display(Name = "Última atualização")]
        public int CotacaoId { get; set; }

        public Cotacao Cotacao { get; set; }

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

        [Display(Name = "Quantidade em unidades de consumo")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float QtdUnddConsumo { get; set; }

        [Display(Name = "Quantidade Múltiplo Compra")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float QtdMltplCompra { get; set; }

        [Display(Name = "Preço Bruto Compra")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float PrcBrtCompra { get; set; }

        [Display(Name = "Cred ICMS")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float CrdtIcms { get; set; }

        [Display(Name = "Cred IPI")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float CrdtIpi { get; set; }

        [Display(Name = "Cred PIS")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float CrdtPis { get; set; }

        [Display(Name = "Cred Cofins")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float CrdtCofins { get; set; }

        [Display(Name = "Soma Cred Impostos")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float SumCrdImpostos { get; set; }

        [Display(Name = "Desp Importação")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float DspImportacao { get; set; }

        [Display(Name = "Custo Extra")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float CustoExtra { get; set; }

        [Display(Name = "Custo")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float Custo { get; set; }

        [Display(Name = "Custo Un Consumo")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float CustoUndCnsm { get; set; }

        [Display(Name = "Pag Forn Import R$/un")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float PgtFornecImp { get; set; }

        [ForeignKey("Produto")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }
    }
}
