using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class ProcTubo
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        public virtual Produto Produto { get; set; }

        [Display(Name = "Resina Base")]
        public int ResinaBaseId { get; set; }

        public virtual ResinaBase ResinaBase { get; set; }

        [Display(Name = "Diam. Externo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DiamExterno { get; set; }

        [Display(Name = "Diam. Interno")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DiamInterno { get; set; }

        [Display(Name = "Série")]
        public int SerieId { get; set; }

        public virtual Serie Serie { get; set; }

        [Display(Name = "Carga 1")]
        public int Carga1Id { get; set; }

        public virtual Carga Carga1 { get; set; }

        [Display(Name = "% Carga 1")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float PctCarga1 { get; set; }

        [Display(Name = "Carga 2")]
        public int Carga2Id { get; set; }

        public virtual Carga Carga2 { get; set; }

        [Display(Name = "% Carga 2")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float PctCarga2 { get; set; }

        [Display(Name = "Sinterizado")]
        public bool Sinterizado { get; set; }

        [Display(Name = "Cod. Resina Adotada")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CodResinaAdotada { get; set; }

        [Display(Name = "RR Max Resina")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RrMaxResina { get; set; }

        [Display(Name = "Bico Ideal")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float BicoIdeal { get; set; }

        [Display(Name = "Mandril Ideal")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float MandrilIdeal { get; set; }

        [Display(Name = "Seção do Extrudado")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float SecaoExtrudado { get; set; }

        [Display(Name = "Perím. / Seção Extrud.")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float PerimSecaoExtrud { get; set; }

        [Display(Name = "Øext final tubo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DiamExtFinalTubo { get; set; }

        [Display(Name = "Øint final tubo")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float DiamIntFinalTubo { get; set; }

        [Display(Name = "Peso Un kg/m Liq")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float PesoUnKgMLiq { get; set; }

        [Display(Name = "ptfe kg/m")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float PtfeKgM { get; set; }

        [Display(Name = "Lubr kg/m")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float LubrKgM { get; set; }

        [Display(Name = "Cod. Preforma Ideal")]
        public int CodPreformaIdeal { get; set; }

        [Display(Name = "RR")]
        public float Rr { get; set; }

        [Display(Name = "Lance Sinterizado")]
        public int LanceSinterizado { get; set; }

        [Display(Name = "Fator Multipl. Quantid.")]
        public int FatorMultiplQtde { get; set; }

        [Display(Name = "Qt. PF")]
        public int QtPf { get; set; }

        [Display(Name = "Embalagem")]
        public int EmbalagemId { get; set; }

        public virtual Embal Embalagem { get; set; }

        [Display(Name = "Quant. por Embalagem")]
        public int QuantEmbalagem { get; set; }

        [Display(Name = "Processo Continuo")]
        public bool ProcessoContinuo { get; set; }

        [Display(Name = "Vextr ÚMIDO m / min")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VextrUmidoMin { get; set; }

        [Display(Name = "Fator Multipl. V.Extr.")]
        public int FatorMultiplVExter { get; set; }

        [Display(Name = "Vsint m/min")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VsintMMin { get; set; }

        [Display(Name = "Fator Multipl. Vel.Sint.")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float FatorMultiplVelSint { get; set; }

        [Display(Name = "V Sint resultante")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VSintResultante { get; set; }

        [Display(Name = "Teste Estq Esto")]
        public bool TesteEstqEsto { get; set; }

        [Display(Name = "Conf.Adt +Dos.Lub")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float ConfAdtDosLub { get; set; }

        [Display(Name = "Peneiramento")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float Peneiramento { get; set; }

        [Display(Name = "Mistura min / m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float MisturaMinM { get; set; }

        [Display(Name = "Preparo Extrus min / m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float PreparoExtrusMinM { get; set; }

        [Display(Name = "Vel Efet Extrusão m/min")]
        [DisplayFormat(DataFormatString = "{0:N1}")]
        public float VelEfetExtrusaoMMin { get; set; }

        [Display(Name = "tu Sinteriz. min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuSinterizadoMinM { get; set; }

        [Display(Name = "tu Produção min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuProducaoMinM { get; set; }

        [Display(Name = "tu insp UDC3 min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuInspUdc3MinM { get; set; }

        [Display(Name = "tu teste estanq, min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuTesteEstanqMinM { get; set; }

        [Display(Name = "tu teste estouro, min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuTesteEstouroMinM { get; set; }

        [Display(Name = "tu Embal. min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuEmbalMinM { get; set; }

        [Display(Name = "TU Total, min/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float TuTotalMinM { get; set; }

        [Display(Name = "Custo PTFE R$/m")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float CustoPtfeRsM { get; set; }

        [Display(Name = "Custo Aditivos R$/m")]
        [DisplayFormat(DataFormatString = "{0:N5}")]
        public float CustoAditivosRsM { get; set; }

        [Display(Name = "Custo Lubrif R$/m")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float CustoLubrifRsM { get; set; }

        [Display(Name = "Custo Embal R$/m")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float CustoEmbalRsM { get; set; }

        [Display(Name = "Custo MOD R$/m")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        public float CustoModRsM { get; set; }

        [Display(Name = "Custo Direto Total R$/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoDiretoTotalRsM { get; set; }

        [Display(Name = "Custo Indireto R$/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoIndiretoRsM { get; set; }

        [Display(Name = "Custo Total R$/m")]
        [DisplayFormat(DataFormatString = "{0:N3}")]
        public float CustoTotalRsM { get; set; }

        [Display(Name = "PV R$/Kg")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvRsKg { get; set; }

        [Display(Name = "Cap. Produção m/h")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CapProducaoMH { get; set; }

        [Display(Name = "Qt. p/ Custo")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtPCusto { get; set; }

        [Display(Name = "PV calculado R$/ m")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvCalculadoRsM { get; set; }

        [StringLength(10)]
        public string Cadastro { get; set; }
    }
}
