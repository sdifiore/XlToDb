using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class CuboEstoque
    {
        public int Id { get; set; }

        [Display(Name = "Codigo")]
        [StringLength(10)]
        public string Apelido { get; set; }

        [Display(Name = "Descrição")]
        [StringLength(256)]
        public string Descricao { get; set; }

        [Display(Name = "Classe de Custo")]
        [StringLength(32)]
        public string ClasseCusto { get; set; }

        [Display(Name = "Quantidade")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public int Quantidade { get; set; }
    }
}
