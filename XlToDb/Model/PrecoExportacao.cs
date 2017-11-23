using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PrecoExportacao
    {
        public int Id { get; set; }

        [StringLength(128)]
        [Display(Name = "Linha / Un")]
        public string LinhaUn { get; set; }

        [StringLength(128)]
        [Display(Name = "Descrição")]
        public string Descricao { get; set; }

        [StringLength(10)]
        [Display(Name = "Código")]
        public string Apelido { get; set; }

        [Display(Name = "Peso Líquido")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float PesoLiquido { get; set; }

        [Display(Name = "Qt/Unid")]
        public int QtUnid { get; set; }

        [Display(Name = "De U$2m a U$5m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float De2A5 { get; set; }

        [Display(Name = "De U$5m a U$10m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float De5A10 { get; set; }

        [Display(Name = "De U$10m a U$20m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float De10A20 { get; set; }

        [Display(Name = "Acima de U$20m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float Acima20 { get; set; }

        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float Com { get; set; }

        [Display(Name = "LL Mínimo")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float LlMin { get; set; }

        [Display(Name = "Condição Preço")]
        public int CondicaoPrecoId { get; set; }

        public CondicaoPreco CondicaoPreco { get; set; }

        [Display(Name = "% Rateio")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PctRateio { get; set; }

        [Display(Name = "Cond Pag Dias")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CondPag { get; set; }

        [Display(Name = "Tx Juros Efetiva")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float IEfetiva { get; set; }

        [Display(Name = "% Espec Frete")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float PctEspecFrete { get; set; }

        [Display(Name = "Desp Exp Padrão")]
        public bool DespExpPadrao { get; set; }

        [Display(Name = "% Desp Export Espec")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float PdtDespExportEspec { get; set; }

        [Display(Name = "PV FOB Maximo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvFobMax { get; set; }

        [Display(Name = "Custo Direto")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoDireto { get; set; }

        [Display(Name = "Rateio Custos Fixos")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float RateioCustoFixo { get; set; }

        [Display(Name = "PV FOB Mínimo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvFobMin { get; set; }

        [Display(Name = "Valor CIF PTFE")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float ValorCifPtfe { get; set; }

        [Display(Name = "Relação PTFE sobre PV")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float RelPtfeSobrePv { get; set; }
    }
}
