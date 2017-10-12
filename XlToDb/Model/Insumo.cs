using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class Insumo
    {
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity), Key]
        public int InsumoId { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }

        public int UnidadeId { get; set; }

        public Unidade Unidade { get; set; }

        public int TipoId { get; set; }

        public Tipo Tipo { get; set; }

        public int ClasseCustoId { get; set; }

        public ClasseCusto ClasseCusto { get; set; }

        public int CategoriaId { get; set; }

        public Categoria Categoria { get; set; }

        public int FamiliaId { get; set; } //

        public Familia Familia { get; set; }

        public int LinhaId { get; set; } //

        public Linha Linha { get; set; }

        [Display(Name = "Peso")]
        public float Peso { get; set; }

        public float QuantidadeCusto { get; set; }

        [Display(Name = "Ativo")]
        public bool Status { get; set; }

        [Display(Name = "Finalidade")]
        public int FinalidadeId { get; set; }

        public Finalidade Finalidade { get; set; }

        [ForeignKey("UnidadeConsumo")]
        [Display(Name = "Unidade de Consumo")]
        public int UnddId { get; set; }

        public Unidade UnidadeConsumo { get; set; }

        [Display(Name = "Quantidade em unidades de consumo")]
        public float QtdUnddConsumo { get; set; }

        [Display(Name = "Quantidade Múltiplo Compra")]
        public float QtdMltplCompra { get; set; }
    }
}
