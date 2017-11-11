using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class MedidaFita
    {
        public int MedidaFitaId { get; set; }

        [Display(Name = "Nome")]
        public string Apelido { get; set; }

        [Display(Name = "Largura")]
        public int LarguraMm { get; set; }

        [Display(Name = "Comprimento")]
        public int ComprimentoMetros { get; set; }
    }
}
