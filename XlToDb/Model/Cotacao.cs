using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Cotacao
    {
        public int CotacaoId { get; set; }

        [Display(Name = "Data")]
        public DateTime DateTime { get; set; }

        [StringLength(10)]
        [Display(Name = "Item")]
        public string Apelido { get; set; }

        [StringLength(256)]
        [Display(Name = "Cotação")]
        public string Descricao { get; set; }

        [ForeignKey("Produto")]
        [Display(Name = "Ítem")]
        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }
    }
}
