using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Operacao
    {
        public int OperacaoId { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string CodigoOperacao { get; set; }

        [StringLength(100)]
        [Display(Name = "Setor de Produção")]
        public string SetorProducao { get; set; }

        [StringLength(100)]
        [Display(Name = "Descrição da Operação")]
        public string Descricao { get; set; }

        [Display(Name = "Taxa de Ocupação")]
        public float TaxaOcupacao { get; set; }

        [StringLength(100)]
        [Display(Name = "Comentário")]
        public string Comentario { get; set; }

        [Display(Name = "Quantidade de Máquinas")]
        public int QtdMaquinas { get; set; }

        [Display(Name = "Custo em R$/h")]
        public float Custo { get; set; }

        [Display(Name = "Setor")]
        public int SetorId { get; set; }

        public Setor Setor { get; set; }
    }
}
