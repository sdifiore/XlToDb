using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class QtdEmbalagem
    {
        public int QtdEmbalagemId { get; set; }

        [Display(Name = "Cartucho Rolo/Cx")]
        public int? CartuchoRolCx { get; set; }

        [Display(Name = "Cartucho Cxs/Palete")]
        public int? CartuchoCxPlt { get; set; }

        [Display(Name = "Display Rolo/Cx")]
        public int? DisplayRolCx { get; set; }

        [Display(Name = "Carretel Rolo/Cx")]
        public int? CarretelRolCx { get; set; }

        [Display(Name = "Carretel Cxs/Palete")]
        public int? CarretelCxPlt { get; set; }

        public int MedidaFitasId { get; set; }

        public MedidaFita MedidaFita { get; set; }
    }
}
