using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Produto
    {
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity), Key]
        public int Id { get; set; }

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

        [Display(Name = "Grupo Rateio")]
        public int GrupoRateioId { get; set; }

        public GrupoRateio GrupoRateio { get; set; }

        [Display(Name = "Pelo Líquido")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float PesoLiquido { get; set; }

        [Display(Name = "Ativo")]
        public bool Ativo { get; set; }

        [Display(Name = "IPI")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float Ipi { get; set; }

        [Display(Name = "Qt/un")]
        public int QtdUnid { get; set; }

        [Display(Name = "Domínio")]
        public int DominioId { get; set; }

        public Dominio Dominio { get; set; }

        [Display(Name = "Tipo Produção")]
        public int TipoProdId { get; set; }

        public virtual TipoProducao TipoProd { get; set; }

        [Display(Name = "PCP")]
        public int PcpId { get; set; }

        public Pcp Pcp { get; set; }

        [Display(Name = "Qt Un por Un Armz")]
        public int QtUnPorUnArmz { get; set; }

        [Display(Name = "Peso Líquido Calculado")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float PesoLiquidoCalc { get; set; }

        [Display(Name = "Item Estrutura")]
        public int ItemStru { get; set; }

        [Display(Name = "Custo Direto Total")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustODirTotal { get; set; }

        [Display(Name = "Custo Mat última etapa")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CstMatUltmEtapa { get; set; }

        [Display(Name = "Custo Mat etapa 1")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CstMatEtapa1 { get; set; }

        [Display(Name = "Custo Mat etapa 2")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CstMatEtapa2 { get; set; }

        [Display(Name = "Custo Mat etapa 3")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CstMatEtapa3 { get; set; }

        [Display(Name = "Custo Tot Materiais")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CstTotMaterial { get; set; }

        [Display(Name = "Custo Direto MOD")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoDirMod { get; set; }

        [Display(Name = "Horas MOD última etapa")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float HorasModUltmEtapa { get; set; }

        [Display(Name = "Horas MOD etapa 1")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float HorasModEtapa1 { get; set; }

        [Display(Name = "Horas MOD etapa 2")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float HorasModEtapa2 { get; set; }

        [Display(Name = "Horas MOD Total")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float HorasModTotal { get; set; }

        [Display(Name = "Capacidade Prod Hora")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float CapProdHora { get; set; }

        [Display(Name = "Lote Mínimo")]
        public int LoteMinimo { get; set; }

        [Display(Name = "Uso Estrutura")]
        public int UsoStru { get; set; }

        [Display(Name = "Custo Dir R$/Kg")]
        public int CustoDir { get; set; }

        [Display(Name = "Relação MOD / Custo Dir")]
        public float RelModCstDir { get; set; }

        [Display(Name = "% mats etapa final")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PctMatEtapaFinal { get; set; }

        [Display(Name = "% mats etapa preced 1")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PctMatEtapa1 { get; set; }

        [Display(Name = "% mats etapa preced 2")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PctMatEtapa2 { get; set; }

        [Display(Name = "% mats etapa preced 3")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PctMatEtapa3 { get; set; }

        public virtual ICollection<Estrutura> Estrutura { get; set; }

        [StringLength(10)]
        public string Input { get; set; }

        [Display(Name = "Custo Fixo Total")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoFixoTotal { get; set; }

        [Display(Name = "MOI Fabricação")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float MoiFabricacao { get; set; }

        [Display(Name = "Outros Custos Fixos Fab")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float OutrosCustosFab { get; set; }

        [Display(Name = "COMACS + COMTEX p/ FPV")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float ComacsComtexFpv { get; set; }

        [Display(Name = "Custo Fixo Admin Logístico FPV")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float CustoFixoAdminLogFpv { get; set; }

        [Display(Name = "R$ MOI + Desp. Fabr / hMOD")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float RsMoiDespFabHMod { get; set; }

        [Display(Name = "R$ SG&A / hMOD")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float RsSgNAHMod { get; set; }

        [Display(Name = "Custo Fixo Total p/ AnR")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CustoFixoTotalAnr { get; set; }

        [Display(Name = "MOI fabricação p/ AnR")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float MoiFabricAnr { get; set; }

        [Display(Name = "Outros custos fabric AnR")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float OutrosCustosFabricAnr { get; set; }

        [Display(Name = "Custo Fixo COMACS + COMTEX AnR")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CustoFixoComacsCmtexAnr { get; set; }

        [Display(Name = "Custo Fixo Admin Logístico Anr")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CustoFixoAdminAnr { get; set; }

        [Display(Name = "Prop Custo Fixo e Total")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PropCustoFixoTotal { get; set; }

        [Display(Name = "Cst dir unidade")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CstDirUnidade { get; set; }

        [Display(Name = "Cst indir unidade")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float CstIndirUnidade { get; set; }

        [Display(Name = "% PTFE em Peso")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PctPtfePeso { get; set; }

        [Display(Name = "Tamanho")]
        public int MedidaFitaId { get; set; }

        public MedidaFita MedidaFita { get; set; }
    }
}
