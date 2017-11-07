using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
        public class AjusteProduto
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }

        [Display(Name = "Custo Fixo Total")]
        public float CustoFixoTotal { get; set; }

        [Display(Name = "MOI Fabricação")]
        public float MoiFabricacao { get; set; }

        [Display(Name = "Outros Custos Fixos Fab")]
        public float OutrosCustosFab { get; set; }

        [Display(Name = "COMACS + COMTEX p/ FPV")]
        public float ComacsComtexFpv { get; set; }

        [Display(Name = "Custo Fixo Admin Logístico FPV")]
        public float CustoFixoAdminFpv { get; set; }

        [Display(Name = "R$ MOI + Desp. Fabr / hMOD")]
        public float RsMoiDespFabHMod { get; set; }

        [Display(Name = "R$ SG&A / hMOD")]
        public float RsSgNAHMod { get; set; }

        [Display(Name = "Custo Fixo Total p/ AnR")]
        public float CustoFixoTotalAnr { get; set; }

        [Display(Name = "MOI fabricação p/ AnR")]
        public float MoiFabricAnr { get; set; }

        [Display(Name = "Outros custos fabric AnR")]
        public float OutrosCustosFabricAnr { get; set; }

        [Display(Name = "Custo Fixo COMACS + COMTEX AnR")]
        public float CustoFixoComacsCmtexAnr { get; set; }

        [Display(Name = "Custo Fixo Admin Logístico Anr")]
        public float CustoFixoAdminAnr { get; set; }

        [Display(Name = "Tamanho")]
        public int MedidaFitaId { get; set; }

        public MedidaFita MedidaFita { get; set; }
    }
}
