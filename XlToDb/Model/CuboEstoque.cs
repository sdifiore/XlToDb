using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class CuboEstoque
    {
        public int Id { get; set; }

        [Display(Name = "Codigo")]
        [StringLength(10)]
        public string Apelido { get; set; }

        [Display(Name = "Quantidade")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public int Quantidade { get; set; }
    }
}
