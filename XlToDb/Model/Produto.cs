using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Produto
    {
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity), Key]
        public int Id { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }

        public float QuantidadeCusto { get; set; }

        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        public int TipoId { get; set; }

        public Tipo Tipo { get; set; }

        public int ClasseCustoId { get; set; }

        public ClasseCusto ClasseCusto { get; set; }

        public int FamiliaId { get; set; }

        public Familia Familia { get; set; }

        public int LinhaId { get; set; }

        public Linha Linha { get; set; }

        public int GrupoRateioId { get; set; }

        public GrupoRateio GrupoRateio { get; set; }

        public int CategoriaId { get; set; }

        public Categoria Categoria { get; set; }
    }
}
