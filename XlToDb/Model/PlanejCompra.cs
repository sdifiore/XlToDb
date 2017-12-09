using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PlanejCompra
    {
        public int Id { get; set; }

        [Display(Name = "Insumo")]
        public int InsumoId { get; set; }

        [Display(Name = "Insumo")]
        public Insumo Insumo { get; set; }

        [Display(Name = "Soma  de 1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float SomaDe1 { get; set; }

        [Display(Name = "Soma  de 2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe2 { get; set; }

        [Display(Name = "Soma  de 3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe3 { get; set; }

        [Display(Name = "Soma  de 4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe4 { get; set; }

        [Display(Name = "Soma  de 5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe5 { get; set; }

        [Display(Name = "Soma  de 6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe6 { get; set; }

        [Display(Name = "Soma  de 7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe7 { get; set; }

        [Display(Name = "Soma  de 8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe8 { get; set; }

        [Display(Name = "Soma  de 9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe9 { get; set; }

        [Display(Name = "Soma  de 10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe10 { get; set; }

        [Display(Name = "Soma  de 11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe11 { get; set; }

        [Display(Name = "Soma  de 12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SomaDe12 { get; set; }

        [Display(Name = "Un Compra")]
        public int UnidadeId { get; set; }

        [Display(Name = "Un Compra")]
        public Unidade Unidade { get; set; }

        [Display(Name = "Fator Consumo")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float FatorConsumo { get; set; }

        [Display(Name = "Lote Compra")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float LoteCompra { get; set; }

        [Display(Name = "Estoque Min")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public int EstoqueMinimo { get; set; }

        [Display(Name = "Preço Unitário Bruto")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float PrecoUnitarioBruto { get; set; }

        [Display(Name = "Custo Unitário")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoUnitario { get; set; }

        [Display(Name = "Crédito Unit ICMS")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CreditoUnitIcms { get; set; }

        [Display(Name = "Crédito Unit IPI")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CreditoUnitIpi { get; set; }

        [Display(Name = "Crédito Unit PIS")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CreditoUnitPis { get; set; }

        [Display(Name = "Crédito Unit COFINS")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CreditoUnitCofins { get; set; }

        [Display(Name = "Pag Fornec Import")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PagFornecImport { get; set; }

        [Display(Name = "II + Desp Import")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float IiDespImport { get; set; }

        [Display(Name = "Mês ref pag 1 fornec")]
        public int MesRefPag1Fornec { get; set; }

        [Display(Name = "Mês ref pag 2 fornec")]
        public int MesRefPag2Fornec { get; set; }

        [Display(Name = "Estoque Inicial")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public int EstoqueInicial { get; set; }

        [Display(Name = "Ano - 11 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos11 { get; set; }

        [Display(Name = "Ano - 10 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos10 { get; set; }

        [Display(Name = "Ano - 9 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos9 { get; set; }

        [Display(Name = "Ano - 8 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos8 { get; set; }

        [Display(Name = "Ano - 7 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos7 { get; set; }

        [Display(Name = "Ano - 6 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos6 { get; set; }

        [Display(Name = "Ano - 5 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos5 { get; set; }

        [Display(Name = "Ano - 4 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos4 { get; set; }

        [Display(Name = "Ano - 3 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos3 { get; set; }

        [Display(Name = "Ano - 2 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos2 { get; set; }

        [Display(Name = "Ano - 1 mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAnoMenos1 { get; set; }

        [Display(Name = "Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float NCMPAno { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float QCMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QCMes12 { get; set; }

        [Display(Name = "Ano - 11 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos11 { get; set; }

        [Display(Name = "Ano - 10 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos10 { get; set; }

        [Display(Name = "Ano - 9 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos9 { get; set; }

        [Display(Name = "Ano - 8 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos8 { get; set; }

        [Display(Name = "Ano - 7 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos7 { get; set; }

        [Display(Name = "Ano - 6 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos6 { get; set; }

        [Display(Name = "Ano - 5 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos5 { get; set; }

        [Display(Name = "Ano - 4 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos4 { get; set; }

        [Display(Name = "Ano - 3 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos3 { get; set; }

        [Display(Name = "Ano - 2 meses")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos2 { get; set; }

        [Display(Name = "Ano - 1 mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAnoMenos1 { get; set; }

        [Display(Name = "Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float SFCMAno { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VBCMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VBCMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float PFNMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFNMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float PFIMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PFIMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float PIIDIMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PIIDIMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VCCMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCCMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VCIcmsMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIcmsMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VCIpiMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float VCIpiMes12 { get; set; }

        [Display(Name = "1° mês")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float CPisCofinsMes1 { get; set; }

        [Display(Name = "2° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes2 { get; set; }

        [Display(Name = "3° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes3 { get; set; }

        [Display(Name = "4° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes4 { get; set; }

        [Display(Name = "5° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes5 { get; set; }

        [Display(Name = "6° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes6 { get; set; }

        [Display(Name = "7° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes7 { get; set; }

        [Display(Name = "8° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes8 { get; set; }

        [Display(Name = "9° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes9 { get; set; }

        [Display(Name = "10° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes10 { get; set; }

        [Display(Name = "11° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes11 { get; set; }

        [Display(Name = "12° mês")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CPisCofinsMes12 { get; set; }

        [Display(Name = "Média Mensal")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float MediaMensal { get; set; }
    }
}
