using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Base
    {
        public int BaseId { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        [Display(Name = "Tipo")]
        public int TipoId { get; set; }

        public Tipo Tipo { get; set; }

        [Display(Name = "Classe de Custo")]
        public int ClasseCustoId { get; set; }

        public ClasseCusto ClasseCusto { get; set; }

        [Display(Name = "Categoria")]
        public int CategoriaId { get; set; }

        public Categoria Categoria { get; set; }

        [Display(Name = "Família")]
        public int FamiliaId { get; set; }

        public Familia Familia { get; set; }

        [Display(Name = "Linha")]
        public int LinhaId { get; set; }

        public Linha Linha { get; set; }
    }
}
