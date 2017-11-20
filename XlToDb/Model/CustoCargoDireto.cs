using System.ComponentModel.DataAnnotations;
using XlToDb.Model;

public class CustoCargoDireto
    {
        public int Id { get; set; }

        [StringLength(32)]
        [Display(Name = "Setor de Produção")]
        public string SetorProducao { get; set; }

        [Display(Name = "Setor")]
        public int SetorId { get; set; }

        public Setor Setor { get; set; }

        [Display(Name = "Operadores")]
        public int Operadores { get; set; }

        
        [Display(Name = "MO Direta")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float MoDireta { get; set; }

        [StringLength(2)]
        [Display(Name = "Código Lider + Apoio")]
        public string CodigoLiderApoio { get; set; }

        [Display(Name = "MO Direta Lideres e Apoio")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float MoDiretaLiderApoio { get; set; }

        [Display(Name = "Horas MOD só operadores")]
        public float HorasModOperadores { get; set; }

        [Display(Name = "Horas MOD total")]
        public float HorasModTotal { get; set; }

        [Display(Name = "Custo unitário")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CustoUnitario { get; set; }

        [Display(Name = "Rateio Setor 20")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RateioSetor20 { get; set; }

        [Display(Name = "Rateio Setor 40")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RateioSetor40 { get; set; }

        [Display(Name = "Rateio Setor 50")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RateioSetor50 { get; set; }

        [Display(Name = "Rateio Setor 60")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RateioSetor60 { get; set; }

        [Display(Name = "Soma dos Indiretos")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaIndiretos { get; set; }

        [Display(Name = "Soma Diretos e Indiretos")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDiretoIndireto { get; set; }

        [Display(Name = "Custo Unitário")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RateioCustoUnitario { get; set; }
    }