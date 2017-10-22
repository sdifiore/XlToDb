using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Estrutura
    {
        public int Id { get; set; }

        [StringLength(10)]
        [Display(Name = "Código do Produto")]
        public string Apelido { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        [Display(Name = "Quantidade para Custo")]
        public float QtdCusto { get; set; }

        public int SequenciaId { get; set; }

        [Display(Name = "Sequência")]
        public Sequencia Sequencia { get; set; }

        [StringLength(10)]
        [Display(Name = "Código Compra")]
        public string Item { get; set; }

        public bool Onera { get; set; }

        [Display(Name = "Lote")]
        public float Lote { get; set; }

        [Display(Name = "Perdas")]
        public float Perda { get; set; }

        [Display(Name = "Observação")]
        public string Observacao { get; set; }

        public int ProdutoId { get; set; }

        public Produto Produto { get; set; }
    }
}
