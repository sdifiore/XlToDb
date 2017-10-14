using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class ParteProduto
    {
        public int ParteProdutoId { get; set; }

        [Display(Name = "Grupo de Rateio")]
        public int GrupoRateioId { get; set; }

        public GrupoRateio GrupoRateio { get; set; }

        [Display(Name = "Peso Líquido")]
        public float Peso { get; set; }

        [Display(Name = "Status")]
        public bool Status { get; set; }

        [Display(Name = "IPI")]
        public float Ipi { get; set; }

        [Display(Name = "Quantidade/Unidade")]
        public int QtdUndd { get; set; }

        [Display(Name = "Domínio")]
        public int DominioId { get; set; }

        public Dominio Dominio { get; set; }

        [Display(Name = "Tipo de Produção")]
        public int TipoProducaoId { get; set; }

        public TipoProducao TipoProducao { get; set; }

        [Display(Name = "PCP")]
        public int PcpId { get; set; }

        public Pcp Pcp { get; set; }

        [Display(Name = "Qtd Unidade por Und Armazenamento")]
        public int QtdUndArmz { get; set; }

        [Display(Name = "Produto")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }
    }
}
