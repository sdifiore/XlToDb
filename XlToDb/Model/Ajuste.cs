using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Ajuste
    {
        public int AjusteId { get; set; }

        [Display(Name = "De")]
        [ForeignKey("Origem")]
        public int OrigemId { get; set; }

        public Produto Origem { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeDeId { get; set; }

        public Unidade UnidadeDe { get; set; }

        [Display(Name = "Para")]
        [ForeignKey("Atual")]
        public int AtualId { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeParaId { get; set; }

        public Unidade UnidadePara { get; set; }

        public Produto Atual { get; set; }

        public float Fator { get; set; }

        public int TipoAlteracaoId { get; set; }

        public TipoAlteracao TipoAlteracao { get; set; }

        public float Medida { get; set; }

    }
}
