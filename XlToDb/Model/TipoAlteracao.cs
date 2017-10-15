using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class TipoAlteracao
    {
        public int TipoAlteracaoId { get; set; }

        [Display(Name = "Alteração")]
        [StringLength(16)]
        public string Descricao { get; set; }
    }
}
