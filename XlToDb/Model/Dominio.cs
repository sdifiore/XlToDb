using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class Dominio
    {
        public int DominioId { get; set; }

        [StringLength(16)]
        [Display(Name = "Domínio da Produção")]
        public string Descricao { get; set; }
    }
}
