using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PlanejNecessid
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int EstruturaId { get; set; }

        [Display(Name = "Código")]
        public Estrutura Estrutura { get; set; }

        [StringLength(128)]
        [Display(Name = "Setor Produção")]
        public string SetorProducao { get; set; }

        [StringLength(128)]
        [Display(Name = "Categoria")]
        public string Categoria { get; set; }

        [StringLength(128)]
        [Display(Name = "Família")]
        public string Familia { get; set; }

        [Display(Name = "Lista Planej Produção")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ListaPlanejProducao { get; set; }

        [Display(Name = "Necessid Compon p/ Produção")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float NeedComponProducao { get; set; }

        [Display(Name = "Lista Necessidade Prod Nivel 1")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ListaNecessProdNivel1 { get; set; }

        [Display(Name = "Nec Compon p/ Lista P-1")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float NecCompListaP1 { get; set; }

        [Display(Name = "Lista Necessidade Prod Nivel 2")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ListaNecessProdNivel2 { get; set; }

        [Display(Name = "Nec Compon p/ Lista P-2")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float NecCompListaP2 { get; set; }

        [Display(Name = "Lista Necessidade Prod Nivel 3")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ListaNecessProdNivel3 { get; set; }

        [Display(Name = "Nec Compon p/ Lista P-3")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float NecCompListaP3 { get; set; }

        [Display(Name = "Lista Necessidade Prod Nivel 4")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float ListaNecessProdNivel4 { get; set; }

        [Display(Name = "Nec Compon p/ Lista P-4")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float NecCompListaP4 { get; set; }

        [Display(Name = "Necessidade Total Componente")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float NecTotalComponente { get; set; }

        [Display(Name = "1º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes1 { get; set; }

        [Display(Name = "2º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes2 { get; set; }

        [Display(Name = "3º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes3 { get; set; }

        [Display(Name = "4º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes4 { get; set; }

        [Display(Name = "5º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes5 { get; set; }

        [Display(Name = "6º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes6 { get; set; }

        [Display(Name = "7º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes7 { get; set; }

        [Display(Name = "8º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes8 { get; set; }

        [Display(Name = "9º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes9 { get; set; }

        [Display(Name = "10º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes10 { get; set; }

        [Display(Name = "11º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes11 { get; set; }

        [Display(Name = "12º Mês")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float Mes12 { get; set; }
    }
}
