using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PlanejMod
    {
        public int Id { get; set; }

        [Display(Name = "Operação")]
        public int OperacaoId { get; set; }

        [Display(Name = "Operação")]
        public Operacao Operacao { get; set; }

        [Display(Name = "Unidade")]
        public int UnidadeId { get; set; }

        [Display(Name = "Unidade")]
        public Unidade Unidade { get; set; }

        [Display(Name = "Soma  de 1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe1 { get; set; }

        [Display(Name = "Soma  de 2° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe2 { get; set; }

        [Display(Name = "Soma  de 3° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe3 { get; set; }

        [Display(Name = "Soma  de 4° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe4 { get; set; }

        [Display(Name = "Soma  de 5° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe5 { get; set; }

        [Display(Name = "Soma  de 6° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe6 { get; set; }

        [Display(Name = "Soma  de 7° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe7 { get; set; }

        [Display(Name = "Soma  de 8° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe8 { get; set; }

        [Display(Name = "Soma  de 0° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe9 { get; set; }

        [Display(Name = "Soma  de 10° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe10 { get; set; }

        [Display(Name = "Soma  de 11° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe11 { get; set; }

        [Display(Name = "Soma  de 12° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe12 { get; set; }

        [Display(Name = "Ano - 11 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos11 { get; set; }

        [Display(Name = "Ano - 10 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos10 { get; set; }

        [Display(Name = "Ano - 9 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos9 { get; set; }

        [Display(Name = "Ano - 8 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos8 { get; set; }

        [Display(Name = "Ano - 7 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos7 { get; set; }

        [Display(Name = "Ano - 6 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos6 { get; set; }

        [Display(Name = "Ano - 5 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos5 { get; set; }

        [Display(Name = "Ano - 4 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos4 { get; set; }

        [Display(Name = "Ano - 3 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos3 { get; set; }

        [Display(Name = "Ano - 2 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos2 { get; set; }

        [Display(Name = "Ano - 1 mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float AnoMenos1 { get; set; }

        [Display(Name = "Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float Ano { get; set; }

        [Display(Name = "Setor")]
        public int SetorId { get; set; }

        [Display(Name = "Setor")]
        public Setor Setor { get; set; }

        [Display(Name = "Média Mensal")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float MediaMensal { get; set; }
    }
}
