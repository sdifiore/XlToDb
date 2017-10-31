using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Estrutura
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }

        [Display(Name = "Unidade Produto")]
        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        [Display(Name = "Quantidade para Custo")]
        public float QtdCusto { get; set; }

        [Display(Name = "Sequência")]
        public int SequenciaId { get; set; }

        public Sequencia Sequencia { get; set; }

        [StringLength(10)]
        [Display(Name = "Código Compra")]
        public string Item { get; set; }

        [StringLength(256)]
        [Display(Name = "Descrição Componente / Processo")]
        public string DescCompProc { get; set; }

        [Display(Name = "Unidade de Compra")]
        public int UnidadeCompraId { get; set; }

        public Unidade UnidadeCompra { get; set; }

        [Display(Name = "Custo unit compra")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CustoUnitCompra { get; set; }

        [Display(Name = "Onera")]
        public bool Onera { get; set; }

        [Display(Name = "Lote")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float Lote { get; set; }

        [DisplayFormat(DataFormatString = "{0:P2}")]
        [Display(Name = "Perdas")]
        public float Perda { get; set; }

        [Display(Name = "Observação")]
        public string Observacao { get; set; }

        [Display(Name = "Parte no Custo")]
        public float PartCusto { get; set; }

        [Display(Name = "Qt Eff unit comp")]
        public float QtEftvUntrCmpnt { get; set; }

        [Display(Name = "Cst Compr Un Prod")]
        public float CstCmprUndPrd { get; set; }

        [Display(Name = "Tipo Item Custo")]
        public string TpItmCst { get; set; }

        [Display(Name = "Alerta SubProduto")]
        public string AlrtSbPrdt { get; set; }

        [Display(Name = "Tempo máquina")]
        public float TempMaq { get; set; }

        [Display(Name = "Qtd por  Unidade")]
        public int QtdUndd { get; set; }

        [Display(Name = "Peso Liq Final")]
        public float PsLiqdFnl { get; set; }

        [Display(Name = "Peso Liq Precedente")]
        public float PsLiqdPrcdt { get; set; }

        [Display(Name = "Horas Mod Final")]
        public float HrsModFnl { get; set; }

        [Display(Name = "Horas Mod Prec 1")]
        public float HrsModPrec1 { get; set; }

        [Display(Name = "Horas Mod Prec 2")]
        public float HrsModPrec2 { get; set; }

        [Display(Name = "Id Produto")]
        public string IdProd { get; set; }

        [Display(Name = "Id Componente")]
        public string IdCmpnt { get; set; }

        [Display(Name = "Prod Horária")]
        public float PdrHoraria { get; set; }

        [Display(Name = "Prod-comp")]
        [StringLength(24)]
        public string ProdComp { get; set; }

        [Display(Name = "Custo Individual")]
        public float CstIndividual { get; set; }

        [Display(Name = "Cust Material Direto")]
        public float CstMtrlDrt { get; set; }

        [Display(Name = "Cust Material Prec 1")]
        public float CstMtrlPrcd1 { get; set; }

        [Display(Name = "Cust Material Prec 2")]
        public float CstMtrlPrcd2 { get; set; }

        [Display(Name = "Cust Material Prec 3")]
        public float CstMtrlPrcd3 { get; set; }

        public bool Header { get; set; }
    }
}
