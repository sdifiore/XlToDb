using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Graxa
    {
        public int Id { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [StringLength(100)]
        [Display(Name = "Descricao")]
        public string Descricao { get; set; }

        [Display(Name = "Embalagem")]
        public int EmbalagemId { get; set; }

        public virtual Embalagem Embalagem { get; set; }

        [Display(Name = "Peso unitário")]
        public float Peso { get; set; }

        [Display(Name = "% Silicone")]
        public float PctSilicone { get; set; }

        [Display(Name = "% Sília")]
        public float PctSilica { get; set; }

        [Display(Name = "% PTFE")]
        public float PctPtfe { get; set; }

        [Display(Name = "Resina")]
        public int ResinaId { get; set; }

        public virtual Resina Resina { get; set; }

        [Display(Name = "Pesagem min/un")]
        public float PesagemMinUn { get; set; }

        [Display(Name = "Mistura min/un")]
        public float Mistura { get; set; }

        [Display(Name = "Embalagem min/un")]
        public float EmbalagemMedida { get; set; }

        [Display(Name = "Rotulagem min/un")]
        public float Rotulagem { get; set; }

        [Display(Name = "Lote mínimo un")]
        public float LoteMinino { get; set; }

        [Display(Name = "PTFE Kg")]
        public float Ptfe { get; set; }

        [Display(Name = "Silicone Kg")]
        public float Silicone { get; set; }

        [Display(Name = "Sílica Kg")]
        public float Silica { get; set; }

        [Display(Name = "Pesagem h")]
        public float PesagemH { get; set; }

        [Display(Name = "Mistura h")]
        public float MisturaH { get; set; }

        [Display(Name = "Embalar h")]
        public float EmbalarH { get; set; }

        [Display(Name = "Rotular h")]
        public float RotularH { get; set; }
    }
}
