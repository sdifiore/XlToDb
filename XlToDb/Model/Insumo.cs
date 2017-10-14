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
        public float PrecoUsd { get; set; }

        [Display(Name = "ICMS")]
        public float Icms { get; set; }

        [Display(Name = "IPI")]
        public float Ipi { get; set; }

        [Display(Name = "PIS")]
        public float Pis { get; set; }

        [Display(Name = "Cofins")]
        public float Cofins { get; set; }

        [Display(Name = "Despesas Extras")]
        public float DespExtra { get; set; }

        [Display(Name = "II + Desp. Importação")]
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
        public float QtdUnddConsumo { get; set; }

        [Display(Name = "Quantidade Múltiplo Compra")]
        public float QtdMltplCompra { get; set; }

        [ForeignKey("Produto")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }
    }
}
