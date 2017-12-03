using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class DfxProdRev
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        [Display(Name = "Quant por Unid")]
        public int QtdUnidade { get; set; }

        [Display(Name = "Quant compra")]
        public int QtdCompra { get; set; }

        [Display(Name = "Preço Compra s/ IPI")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float PrecoCompra { get; set; }

        [Display(Name = "Valor Compra")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ValorCompra { get; set; }

        [Display(Name = "Rateio DFix por Produto")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float RateioDfixProduto { get; set; }

        [Display(Name = "Rateio DFix por Unidade")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float RateioDfixUnidade { get; set; }

        [Display(Name = "MOI Fábrica")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float MoiFabrica { get; set; }

        [Display(Name = "Desps Fábrica")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DespsFabrica { get; set; }

        [Display(Name = "Desps Dep Comercial")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DespsDepComercial { get; set; }

        [Display(Name = "Desps Depts Admin-Log")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DespsDeptAdminLog { get; set; }

        [Display(Name = "Quant Rolos")]
        public int QtdRolos { get; set; }

        [Display(Name = "Rateio DFix por Produto unitário")]
        public float RateioDfixProdutoUnitario { get; set; }

        [Display(Name = "Proporção Custo")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float ProporcaoCusto { get; set; }

        [Display(Name = "Proporção em cxs")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float ProporcaoEmCxs { get; set; }
    }
}
