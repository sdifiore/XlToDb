using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PreForma
    {
        public int Id { get; set; }

        [Display(Name = "Pré Forma n°")]
        public int PreFormaNum { get; set; }

        [Display(Name = "Forma - ØE mm")]
        public int FormaDiamE { get; set; }

        [Display(Name = "Vareta - ØI mm")]
        public int VaretaDiamI { get; set; }

        [StringLength(8)]
        public string Medidas { get; set; }

        [Display(Name = "Comprimento mm")]
        public int Comprimento { get; set; }

        [Display(Name = "T.U.P. min")]
        public int Tup { get; set; }

        [Display(Name = "Prensa Pré-Forma")]
        public int PrensaPreFormaId { get; set; }

        public PrensaPreForma PrensaPreForma { get; set; }

        [Display(Name = "Preparo (min)")]
        public int Preparo { get; set; }

        [Display(Name = "Troca PF (min)")]
        public float TrocaPf { get; set; }

        [Display(Name = "Extrusora")]
        public int ExtrusoraId { get; set; }

        public Extrusora Extrusora { get; set; }

        [Display(Name = "Seçãoda PF")]
        public float SecaoPf { get; set; }

        [Display(Name = "Kgf Prensagem")]
        public float KgfPrensagem { get; set; }

        [Display(Name = "Pressão Óleo")]
        public float PressaoOleo { get; set; }

        [Display(Name = "Ø Pistão Hidráulico")]
        public float DiamPistaoHidraulico { get; set; }

        [Display(Name = "Kg PF úmida")]
        public float KgPfUmida { get; set; }
    }
}
