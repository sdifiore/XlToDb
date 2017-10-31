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

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        [Display(Name = "Tipo")]
        public int TipoId { get; set; }

        public Tipo Tipo { get; set; }

        [Display(Name = "Classe de Custo")]
        public int ClasseCustoId { get; set; }

        public ClasseCusto ClasseCusto { get; set; }

        [Display(Name = "Categoria")]
        public int CategoriaId { get; set; }

        public Categoria Categoria { get; set; }

        [Display(Name = "Família")]
        public int FamiliaId { get; set; }

        public Familia Familia { get; set; }

        [Display(Name = "Linha")]
        public int LinhaId { get; set; }

        public Linha Linha { get; set; }

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

        [Display(Name = "Preço Bruto Compra")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PrcBrtCompra { get; set; }

        [Display(Name = "Cred ICMS")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CrdtIcms { get; set; }

        [Display(Name = "Cred IPI")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CrdtIpi { get; set; }

        [Display(Name = "Cred PIS")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CrdtPis { get; set; }

        [Display(Name = "Cred Cofins")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CrdtCofins { get; set; }

        [Display(Name = "Soma Cred Impostos")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float SumCrdImpostos { get; set; }

        [Display(Name = "Desp Importação")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DspImportacao { get; set; }

        [Display(Name = "Custo Extra")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CustoExtra { get; set; }

        [Display(Name = "Custo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Custo { get; set; }

        [Display(Name = "Custo Un Consumo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CustoUndCnsm { get; set; }

        [Display(Name = "Pag Forn Import R$/un")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PgtFornecImp { get; set; }

        [Display(Name = "Uso em estruturas")]
        public int UsoStru { get; set; }
    }
}
