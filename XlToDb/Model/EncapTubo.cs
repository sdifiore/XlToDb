using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class EncapTubo
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        [Display(Name = "Dext revest")]
        public float DextRevest { get; set; }

        [Display(Name = "Dint revest")]
        public float DintRevest { get; set; }

        [StringLength(10)]
        [Display(Name = "Resina Base")]
        public string ResinaBase { get; set; }

        [StringLength(10)]
        [Display(Name = "Aditivo")]
        public string Aditivo { get; set; }

        [Display(Name = "Den Rev")]
        public float DenRev { get; set; }

        [Display(Name = "Peso Revest")]
        public float PesoRevest { get; set; }

        [Display(Name = "Vel Revest m/min")]
        public int VelRevest { get; set; }

        [Display(Name = "% carga")]
        [DisplayFormat(DataFormatString = "{0:P}")]
        public float PctCarga { get; set; }

        [Display(Name = "Resina Kg/m")]
        public float Resina { get; set; }

        [Display(Name = "Master Kg/m")]
        public float Master { get; set; }
    }
}
