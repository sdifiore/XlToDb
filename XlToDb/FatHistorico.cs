using System;
using System.ComponentModel.DataAnnotations;
using XlToDb.Model;

namespace XlToDb
{
    public class FatHistorico
    {
        public int Id { get; set; }

        [StringLength(32)]
        [Display(Name = "Número Pedido")]
        public string NumeroPedido { get; set; }

        [DataType(DataType.Date)]
        [Display(Name = "Data Pedido")]
        public DateTime DataPedido { get; set; }

        [StringLength(256)]
        [Display(Name = "Cliente")]
        public string Cliente { get; set; }

        [StringLength(64)]
        [Display(Name = "Estado")]
        public string Estado { get; set; }

        [StringLength(64)]
        [Display(Name = "Cidade")]
        public string Cidade { get; set; }

        [StringLength(64)]
        [Display(Name = "Região")]
        public string Regiao { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        [Display(Name = "Código")]
        public Produto Produto { get; set; }

        [StringLength(256)]
        [Display(Name = "Vendedor")]
        public string Vendedor { get; set; }

        [StringLength(64)]
        [Display(Name = "Forma de Pagamento")]
        public string FormaPagamento { get; set; }

        [DataType(DataType.Date)]
        [Display(Name = "Data de Validade")]
        public DateTime DataValidade { get; set; }

        [StringLength(16)]
        [Display(Name = "Tipo Venda")]
        public string TipoVenda { get; set; }

        [DataType(DataType.Date)]
        [Display(Name = "Data Faturamento")]
        public DateTime DataFaturamento { get; set; }

        [Display(Name = "Nota Fiscal")]
        public int NotaFiscal { get; set; }

        [Display(Name = "Quantidade")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float Quantidade { get; set; }

        [Display(Name = "Valor Unitário")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ValorUnitario { get; set; }

        [Display(Name = "Valor Mercadoria")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ValorMercadoria { get; set; }

        [Display(Name = "Valor IPI")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ValorIpi { get; set; }

        [Display(Name = "Valor Subst Tributária")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ValorSubstTributaria { get; set; }

        [Display(Name = "Prazo Médio Recebimento")]
        public int PrazoMedioRecebimento { get; set; }

        [Display(Name = "Rec Bruta")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float RecBruta { get; set; }

        [Display(Name = "Fatur Bruto")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float FaturBruto { get; set; }

        [StringLength(64)]
        [Display(Name = "Mês Cadastro")]
        public string MesCadastro { get; set; }

        [StringLength(16)]
        [Display(Name = "Ano Mês Fatura")]
        public string AnoMesFatura { get; set; }

        [StringLength(64)]
        [Display(Name = "Situação")]
        public string Situacao { get; set; }

        [StringLength(64)]
        [Display(Name = "Mês Entrega")]
        public string MesEntrega { get; set; }

        [StringLength(1024)]
        [Display(Name = "Cliente-Pedido")]
        public string ClientePedido { get; set; }

        [Display(Name = "Cod Produto Ajustado")]
        public int ProdutoAjustadoId { get; set; }

        [Display(Name = "Cod Produto Ajustado")]
        public Produto ProdutoAjustado { get; set; }

        [StringLength(8)]
        [Display(Name = "Natur Cli")]
        public string NaturCli { get; set; }

        [Display(Name = "Peso Produto")]
        public float PesoProduto { get; set; }

        [StringLength(64)]
        [Display(Name = "Tipo Cliente")]
        public string TipoCliente { get; set; }

        [StringLength(128)]
        [Display(Name = "Categoria Cliente")]
        public string CategoriaCliente { get; set; }

        [StringLength(128)]
        [Display(Name = "Segmento Cliente")]
        public string SegmentoCliente { get; set; }

        [StringLength(512)]
        [Display(Name = "Grupo")]
        public string Grupo { get; set; }

        [Display(Name = "Prazo Entrega")]
        public int PrazoEntrega { get; set; }

        [Display(Name = "Tx Dólar")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float TxDolar { get; set; }

        [Display(Name = "Fat Bruto USD")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float FatBrutoUsd { get; set; }

        [Display(Name = "Preço Individual")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PrecoIndividual { get; set; }

        [Display(Name = "Receita Líquida")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ReceitaLiquida { get; set; }

        [Display(Name = "Comissão")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Comissao { get; set; }

        [Display(Name = "Frete")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Frete { get; set; }

        [Display(Name = "Cst Financ + Cobrança")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CstFinBobranca { get; set; }

        [Display(Name = "Quant Ajustada")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float QuantAjustada { get; set; }

        [Display(Name = "ICMS")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Icms { get; set; }

        [Display(Name = "Prazo x Fatur")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PrazoFatur { get; set; }

        [Display(Name = "Horas MOD")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float HorasMod { get; set; }

        [Display(Name = "Com GV Comacs")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ComGvComacs { get; set; }

        [StringLength(512)]
        [Display(Name = "Descr Prod Ajustado")]
        public string DescrProdAjustado { get; set; }

        [StringLength(512)]
        [Display(Name = "Prod Categoria Ajustado")]
        public string ProdCategoriaAjustado { get; set; }
    }
}
