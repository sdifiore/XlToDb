using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Estrutura
    {
        public int Id { get; set; }

        [StringLength(10)]
        public string Apelido { get; set; }

        public int ProdutoId { get; set; }

        [ForeignKey("ProdutoId")]
        public Produto Produto { get; set; }

        public bool Onera { get; set; }

        public string Observacao { get; set; }

        public int SequenciaId { get; set; }

        public Sequencia Sequencia { get; set; }
    }
}
