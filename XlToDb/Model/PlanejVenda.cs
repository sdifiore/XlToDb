using System;
using System.ComponentModel.DataAnnotations;

namespace XlToDb.Model
{
    public class PlanejVenda
    {
        public int Id { get; set; }

        [Display(Name = "Código")]
        public int ProdutoId { get; set; }

        [Display(Name = "Código")]
        public Produto Produto { get; set; }

        [Display(Name = "Peso Unitário")]
        public float PesoUnitario { get; set; }

        [Display(Name = "Horas MOD")]
        public float HorasMod { get; set; }

        [Display(Name = "Custo Direto Total")]
        public float CustoDiretoTotal { get; set; }

        [Display(Name = "Custo Direto só Mats")]
        public float CustoDiretoMats { get; set; }

        [Display(Name = "Custo Direto só MOD")]
        public float CustoDiretoMod { get; set; }

        [Display(Name = "Custo Fixo Fábrica")]
        public float CustoFixoFabrica { get; set; }

        [Display(Name = "Custo Fixo Adm + Com")]
        public float CustoFixAdmCom { get; set; }

        [Display(Name = "Alíquota IPI")]
        public float AliquotaIpi { get; set; }

        [Display(Name = "QvQtNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos12 { get; set; }

        [Display(Name = "QvQtNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos11 { get; set; }

        [Display(Name = "QvQtNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos10 { get; set; }

        [Display(Name = "QvQtNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos9 { get; set; }

        [Display(Name = "QvQtNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos8 { get; set; }

        [Display(Name = "QvQtNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos7 { get; set; }

        [Display(Name = "QvQtNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos6 { get; set; }

        [Display(Name = "QvQtNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos5 { get; set; }

        [Display(Name = "QvQtNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos4 { get; set; }

        [Display(Name = "QvQtNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos3 { get; set; }

        [Display(Name = "QvQtNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAnoMenos2 { get; set; }

        [Display(Name = "QvQtNac Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QvQtNacAno { get; set; }

        public DateTime? RefAno { get; set; }

        [Display(Name = "Média Mensal")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtNacMediaMensal { get; set; }

        [Display(Name = "PV med 1º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMed1o3m { get; set; }

        [Display(Name = "PV med 2º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMed2o3m { get; set; }

        [Display(Name = "PV med 3º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMed3o3m { get; set; }

        [Display(Name = "PV med 4º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMed4o3m { get; set; }

        [Display(Name = "PV Nac adotado")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvNacAdotado { get; set; }

        [Display(Name = "ST Média")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float StMedia { get; set; }

        [Display(Name = "ICMS Médio")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float IcmsMedio { get; set; }

        [Display(Name = "Comissão Média Nac")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float ComissaoMediaNac { get; set; }

        [Display(Name = "Fretes Nac %a")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float FreteNacPct { get; set; }

        [Display(Name = "Mês Receb Med Nac")]       // AM
        public int MesRecebMedNac { get; set; }

        [Display(Name = "QtExp Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos12 { get; set; }

        [Display(Name = "QtExp Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos11 { get; set; }

        [Display(Name = "QtExp Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos10 { get; set; }

        [Display(Name = "QtExp Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos9 { get; set; }

        [Display(Name = "QtExp Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos8 { get; set; }

        [Display(Name = "QtExp Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos7 { get; set; }

        [Display(Name = "QtExp Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos6 { get; set; }

        [Display(Name = "QtExp Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos5 { get; set; }

        [Display(Name = "QtExp Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos4 { get; set; }

        [Display(Name = "QtExp Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos3 { get; set; }

        [Display(Name = "QtExp Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAnoMenos2 { get; set; }

        [Display(Name = "QtExp Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpAno { get; set; }

        [Display(Name = "Média Mensal")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtExpMediaMensal { get; set; }

        [Display(Name = "PV med ex 1º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMedEx1o3m { get; set; }

        [Display(Name = "PV med ex 2º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMedEx2o3m { get; set; }

        [Display(Name = "PV med ex 3º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMedEx3o3m { get; set; }

        [Display(Name = "PV med ex 4º 3M")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMedEx4o3m { get; set; }

        [Display(Name = "PV ex adotado")]       // BE
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public float PvMedExAdotado { get; set; }

        [Display(Name = "Comissão Exp %")]
        [DisplayFormat(DataFormatString = "{0:P2}")]
        public float ComissaoExpPct { get; set; }

        [Display(Name = "Prazo Receb Med Exp")]
        public int PrazoRecebMedExp { get; set; }

        [StringLength(128)]      // BI
        [Display(Name = "Comentário")]
        public string ComentarioCelulaBi { get; set; }

        [StringLength(1)]      // BJ
        [Display(Name = "Critério")]
        public string Criterio { get; set; }

        [Display(Name = "Var TC1")]
        public int VartC1 { get; set; }

        [Display(Name = "Var TC2")]
        public int VarTc2 { get; set; }

        [Display(Name = "Var TC3")]
        public int VartC3 { get; set; }

        [Display(Name = "Var TC4")]
        public int VartC4 { get; set; }

        [Display(Name = "PqQtNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos12 { get; set; }

        [Display(Name = "PqQtNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos11 { get; set; }

        [Display(Name = "PqQtNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos10 { get; set; }

        [Display(Name = "PqQtNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos9 { get; set; }

        [Display(Name = "PqQtNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos8 { get; set; }

        [Display(Name = "PqQtNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos7 { get; set; }

        [Display(Name = "PqQtNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos6 { get; set; }

        [Display(Name = "PqQtNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos5 { get; set; }

        [Display(Name = "PqQtNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos4 { get; set; }

        [Display(Name = "PqQtNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos3 { get; set; }

        [Display(Name = "PqQtNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAnoMenos2 { get; set; }

        [Display(Name = "PqQtNac Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacAno { get; set; }

        [Display(Name = "PqQtNac Total")]       // CA
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqQtNacTotal { get; set; }

        [Display(Name = "PplKgNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos12 { get; set; }

        [Display(Name = "PplKgNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos11 { get; set; }

        [Display(Name = "PplKgNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos10 { get; set; }

        [Display(Name = "PplKgNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos9 { get; set; }

        [Display(Name = "PplKgNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos8 { get; set; }

        [Display(Name = "PplKgNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos7 { get; set; }

        [Display(Name = "PplKgNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos6 { get; set; }

        [Display(Name = "PplKgNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos5 { get; set; }

        [Display(Name = "PplKgNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos4 { get; set; }

        [Display(Name = "PplKgNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos3 { get; set; }

        [Display(Name = "PplKgNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAnoMenos2 { get; set; }

        [Display(Name = "PplKgNac Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacAno { get; set; }

        [Display(Name = "PplKgNac Total")]      // CN
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PplKgNacTotal { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos12 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos11 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos10 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos9 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos8 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos7 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos6 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos5 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos4 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos3 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAnoMenos2 { get; set; }

        [Display(Name = "PvvpvaVarPv Ano")]     // CZ
        [DisplayFormat(DataFormatString = "{0:P0}")]
        public float PvvpvaVarPvAno { get; set; }

        [Display(Name = "PcbRbNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos12 { get; set; }

        [Display(Name = "PcbRbNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos11 { get; set; }

        [Display(Name = "PcbRbNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos10 { get; set; }

        [Display(Name = "PcbRbNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos9 { get; set; }

        [Display(Name = "PcbRbNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos8 { get; set; }

        [Display(Name = "PcbRbNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos7 { get; set; }

        [Display(Name = "PcbRbNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos6 { get; set; }

        [Display(Name = "PcbRbNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos5 { get; set; }

        [Display(Name = "PcbRbNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos4 { get; set; }

        [Display(Name = "PcbRbNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos3 { get; set; }

        [Display(Name = "PcbRbNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAnoMenos2 { get; set; }

        [Display(Name = "PcbRbNac Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacAno { get; set; }

        [Display(Name = "RecBruta Nac")]        // DM
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PcbRbNacTotal { get; set; }

        [Display(Name = "PipiIpiNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos12 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos11 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos10 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos9 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos8 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos7 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos6 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos5 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos4 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos3 { get; set; }

        [Display(Name = "PipiIpiNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAnoMenos2 { get; set; }

        [Display(Name = "PipiIpiNac Ano")]      // DY
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PipiIpiNacAno { get; set; }

        [Display(Name = "PstStNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos12 { get; set; }

        [Display(Name = "PstStNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos11 { get; set; }

        [Display(Name = "PstStNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos10 { get; set; }

        [Display(Name = "PstStNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos9 { get; set; }

        [Display(Name = "PstStNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos8 { get; set; }

        [Display(Name = "PstStNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos7 { get; set; }

        [Display(Name = "PstStNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos6 { get; set; }

        [Display(Name = "PstStNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos5 { get; set; }

        [Display(Name = "PstStNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos4 { get; set; }

        [Display(Name = "PstStNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos3 { get; set; }

        [Display(Name = "PstStNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAnoMenos2 { get; set; }

        [Display(Name = "PstStNac Ano")]        // EK
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PstStNacAno { get; set; }

        [Display(Name = "PfbFatBr Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos12 { get; set; }

        [Display(Name = "PfbFatBr Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos11 { get; set; }

        [Display(Name = "PfbFatBr Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos10 { get; set; }

        [Display(Name = "PfbFatBr Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos9 { get; set; }

        [Display(Name = "PfbFatBr Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos8 { get; set; }

        [Display(Name = "PfbFatBr Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos7 { get; set; }

        [Display(Name = "PfbFatBr Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos6 { get; set; }

        [Display(Name = "PfbFatBr Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos5 { get; set; }

        [Display(Name = "PfbFatBr Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos4 { get; set; }

        [Display(Name = "PfbFatBr Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos3 { get; set; }

        [Display(Name = "PfbFatBr Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAnoMenos2 { get; set; }

        [Display(Name = "PfbFatBr Ano")]        // EW
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PfbFatBrAno { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos12 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos11 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos10 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos9 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos8 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos7 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos6 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos5 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos4 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos3 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAnoMenos2 { get; set; }

        [Display(Name = "PicmsIcmsNac Ano")]        // FI
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PicmsIcmsNacAno { get; set; }

        [Display(Name = "Fat Bruto Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnFatBrutoNac { get; set; }

        [Display(Name = "Impostos")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnImpostos { get; set; }

        [Display(Name = "Rec Liquida Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnRecLiquidaNacional { get; set; }

        [Display(Name = "Comissão Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnComissaoNac { get; set; }

        [Display(Name = "FretesNac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnFreteNac { get; set; }

        [Display(Name = "Rec Líquita de Venda Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnRecLiquidaVendaNac { get; set; }

        [Display(Name = "Custo Direto Materiais Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnCustoDirMateriaisNac { get; set; }

        [Display(Name = "Custo Direto de MOD Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnCustoDirModNac { get; set; }

        [Display(Name = "Custo Fixo Fábrica Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnFixoFabricaNac { get; set; }

        [Display(Name = "M.C. Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnMCNac { get; set; }

        [Display(Name = "M.C. Nac %")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float CrnMCNacPct { get; set; }

        [Display(Name = "Custo Fixo Adm + Com Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnCustoFixoAdmComNac { get; set; }

        [Display(Name = "Resultado Bruto Nac")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnResultadoBrutoNac { get; set; }

        [Display(Name = "Resultado Bruto Nac %")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CrnResultadoBrutoNacPct { get; set; }

        [StringLength(128)]
        [Display(Name = "Comentário")]
        public string CrnComentario { get; set; }

        [StringLength(1)]
        [Display(Name = "Critério")]
        public string PqeCriterio { get; set; }

        [Display(Name = "aum-dim")]
        public int PqeAumDim { get; set; }

        [Display(Name = "PqeQtExp Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos12 { get; set; }

        [Display(Name = "PqeQtExp Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos11 { get; set; }

        [Display(Name = "PqeQtExp Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos10 { get; set; }

        [Display(Name = "PqeQtExp Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos9 { get; set; }

        [Display(Name = "PqeQtExp Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos8 { get; set; }

        [Display(Name = "PqeQtExp Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos7 { get; set; }

        [Display(Name = "PqeQtExp Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos6 { get; set; }

        [Display(Name = "PqeQtExp Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos5 { get; set; }

        [Display(Name = "PqeQtExp Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos4 { get; set; }

        [Display(Name = "PqeQtExp Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos3 { get; set; }

        [Display(Name = "PqeQtExp Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAnoMenos2 { get; set; }

        [Display(Name = "PqeQtExp Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpAno { get; set; }

        [Display(Name = "PqeQtExp Total")]      // GM
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PqeQtExpTotal { get; set; }

        [Display(Name = "PpleKgExp Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos12 { get; set; }

        [Display(Name = "PpleKgExp Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos11 { get; set; }

        [Display(Name = "PpleKgExp Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos10 { get; set; }

        [Display(Name = "PpleKgExp Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos9 { get; set; }

        [Display(Name = "PpleKgExp Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos8 { get; set; }

        [Display(Name = "PpleKgExp Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos7 { get; set; }

        [Display(Name = "PpleKgExp Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos6 { get; set; }

        [Display(Name = "PpleKgExp Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos5 { get; set; }

        [Display(Name = "PpleKgExp Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos4 { get; set; }

        [Display(Name = "PpleKgExp Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos3 { get; set; }

        [Display(Name = "PpleKgExp Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAnoMenos2 { get; set; }

        [Display(Name = "PpleKgExp Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpAno { get; set; }

        [Display(Name = "PpleKgExp Total")]     // GZ
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PpleKgExpTotal { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos12 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos11 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos10 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos9 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos8 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos7 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos6 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos5 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos4 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos3 { get; set; }

        [Display(Name = "PvppvaVPVex Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAnoMenos2 { get; set; }

        [Display(Name = "PvppvaVPVex Ano")]     // HL
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PvppvaVPVexAno { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos12 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos11 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos10 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos9 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos8 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos7 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos6 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos5 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos4 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos3 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAnoMenos2 { get; set; }

        [Display(Name = "PreUsdRcExpUs Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsAno { get; set; }

        [Display(Name = "PreUsdRcExpUs Rec Export")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float PreUsdRcExpUsRecExportUSD { get; set; }

        [Display(Name = "PreUsdRcExpUs Rec Export R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreRecExportRs { get; set; }

        [Display(Name = "Comissão Export R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreComissaoExportRs { get; set; }

        [Display(Name = "Frete Desp Export R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreFreteDespExportRs { get; set; }

        [Display(Name = "Rec Liq Venda Export R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreRecLiqVendaExportRs { get; set; }

        [Display(Name = "Custo Direto Mat + MOD Exp R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreCustoDiretoMatModExptRs { get; set; }

        [Display(Name = "Custo Fixo Fab Exp R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreCustoFixoFabExpRs { get; set; }

        [Display(Name = "M.C. Export R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreMCExportRs { get; set; }

        [Display(Name = "M.C. Export %")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float CreMCExportPct { get; set; }

        [Display(Name = "Custo Fixo Adm Com Exp R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreCustoFixoAdmComExpRs { get; set; }

        [Display(Name = "Resultado Bruto Exp R$")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CreResultadoBrutoExpRs { get; set; }

        [Display(Name = "Res Bruto Exp %")]     // IJ
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float CreResBrutoExpPct { get; set; }

        [Display(Name = "Quant Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FbQuantTotal { get; set; }

        [Display(Name = "Peso Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FbPesoTotal { get; set; }

        [Display(Name = "Faturamento Bruto Total")]     // IM
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FbFaturamentoBrutoTotal { get; set; }

        [Display(Name = "Impostos Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RlImpostoTotal { get; set; }

        [Display(Name = "Receita Líquida Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RlReceitaLiquidaTotal { get; set; }

        [Display(Name = "Comissões Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RlComissaoTotal { get; set; }

        [Display(Name = "Fretes + Desp. Exp Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RlFreteDespExpTotal { get; set; }

        [Display(Name = "Receita Liquida Vendas")]      // IR
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RlRecLiqVendaTotal { get; set; }

        [Display(Name = "Custo Dir Mat MOD Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float McCustoDirMatModTotal { get; set; }

        [Display(Name = "Custo Fixo Fábrica Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float McCustoFixoFabricaTotal { get; set; }

        [Display(Name = "Margem Contrib Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float McMargemContribTotal { get; set; }

        [Display(Name = "M.C. Total %")]        // IW
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float McMCbTotalPct { get; set; }

        [Display(Name = "Custo Fixo Com Adm Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RoCustoFixoComAdmTotal { get; set; }

        [Display(Name = "Custo Fixo Com Adm Ajustado Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RoCustoFixoComAdmAjustadoTotal { get; set; }

        [Display(Name = "Resultado Operacional Total")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RoResultadoOperacionalTotal { get; set; }

        [Display(Name = "Resultado Operacional Total %")]
        [DisplayFormat(DataFormatString = "{0:P1}")]
        public float RoResultadoOperacionalTotalPct { get; set; }

        [Display(Name = "CdTot Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos12 { get; set; }

        [Display(Name = "CdTot Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos11 { get; set; }

        [Display(Name = "CdTot Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos10 { get; set; }

        [Display(Name = "CdTot Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos9 { get; set; }

        [Display(Name = "CdTot Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos8 { get; set; }

        [Display(Name = "CdTot Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos7 { get; set; }

        [Display(Name = "CdTot Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos6 { get; set; }

        [Display(Name = "CdTot Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos5 { get; set; }

        [Display(Name = "CdTot Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos4 { get; set; }

        [Display(Name = "CdTot Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos3 { get; set; }

        [Display(Name = "CdTot Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAnoMenos2 { get; set; }

        [Display(Name = "CdTot Ano")]       // JM
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdTotAno { get; set; }

        [Display(Name = "GifTot Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos12 { get; set; }

        [Display(Name = "GifTot Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos11 { get; set; }

        [Display(Name = "GifTot Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos10 { get; set; }

        [Display(Name = "GifTot Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos9 { get; set; }

        [Display(Name = "GifTot Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos8 { get; set; }

        [Display(Name = "GifTot Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos7 { get; set; }

        [Display(Name = "GifTot Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos6 { get; set; }

        [Display(Name = "GifTot Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos5 { get; set; }

        [Display(Name = "GifTot Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos4 { get; set; }

        [Display(Name = "GifTot Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos3 { get; set; }

        [Display(Name = "GifTot Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAnoMenos2 { get; set; }

        [Display(Name = "GifTot Ano")]      // JY
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float GifTotAno { get; set; }

        [Display(Name = "ComTot Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos12 { get; set; }

        [Display(Name = "ComTot Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos11 { get; set; }

        [Display(Name = "ComTot Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos10 { get; set; }

        [Display(Name = "ComTot Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos9 { get; set; }

        [Display(Name = "ComTot Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos8 { get; set; }

        [Display(Name = "ComTot Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos7 { get; set; }

        [Display(Name = "ComTot Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos6 { get; set; }

        [Display(Name = "ComTot Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos5 { get; set; }

        [Display(Name = "ComTot Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos4 { get; set; }

        [Display(Name = "ComTot Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos3 { get; set; }

        [Display(Name = "ComTot Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAnoMenos2 { get; set; }

        [Display(Name = "ComTot Ano")]      // KK
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float ComTotAno { get; set; }

        [Display(Name = "FrtNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos12 { get; set; }

        [Display(Name = "FrtNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos11 { get; set; }

        [Display(Name = "FrtNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos10 { get; set; }

        [Display(Name = "FrtNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos9 { get; set; }

        [Display(Name = "FrtNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos8 { get; set; }

        [Display(Name = "FrtNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos7 { get; set; }

        [Display(Name = "FrtNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos6 { get; set; }

        [Display(Name = "FrtNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos5 { get; set; }

        [Display(Name = "FrtNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos4 { get; set; }

        [Display(Name = "FrtNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos3 { get; set; }

        [Display(Name = "FrtNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAnoMenos2 { get; set; }

        [Display(Name = "FrtNac Ano")]      // KW
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrtNacAno { get; set; }

        [Display(Name = "FrDexp Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos12 { get; set; }

        [Display(Name = "FrDexp Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos11 { get; set; }

        [Display(Name = "FrDexp Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos10 { get; set; }

        [Display(Name = "FrDexp Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos9 { get; set; }

        [Display(Name = "FrDexp Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos8 { get; set; }

        [Display(Name = "FrDexp Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos7 { get; set; }

        [Display(Name = "FrDexp Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos6 { get; set; }

        [Display(Name = "FrDexp Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos5 { get; set; }

        [Display(Name = "FrDexp Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos4 { get; set; }

        [Display(Name = "FrDexp Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos3 { get; set; }

        [Display(Name = "FrDexp Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAnoMenos2 { get; set; }

        [Display(Name = "FrDexp Ano")]      // LI
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float FrDexpAno { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos12 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos11 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos10 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos9 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos8 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos7 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos6 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos5 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos4 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos3 { get; set; }

        [Display(Name = "CdsmtCfMat Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAnoMenos2 { get; set; }

        [Display(Name = "CdsmtCfMat Ano")]      // LU
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float CdsmtCfMatAno { get; set; }

        [Display(Name = "QtvQtTott Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos12 { get; set; }

        [Display(Name = "QtvQtTott Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos11 { get; set; }

        [Display(Name = "QtvQtTott Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos10 { get; set; }

        [Display(Name = "QtvQtTott Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos9 { get; set; }

        [Display(Name = "QtvQtTott Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos8 { get; set; }

        [Display(Name = "QtvQtTott Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos7 { get; set; }

        [Display(Name = "QtvQtTott Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos6 { get; set; }

        [Display(Name = "QtvQtTott Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos5 { get; set; }

        [Display(Name = "QtvQtTott Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos4 { get; set; }

        [Display(Name = "QtvQtTott Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos3 { get; set; }

        [Display(Name = "QtvQtTott Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAnoMenos2 { get; set; }

        [Display(Name = "QtvQtTott Ano")]       // MG
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float QtvQtTottAno { get; set; }

        [Display(Name = "HorasProd Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos12 { get; set; }

        [Display(Name = "HorasProd Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos11 { get; set; }

        [Display(Name = "HorasProd Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos10 { get; set; }

        [Display(Name = "HorasProd Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos9 { get; set; }

        [Display(Name = "HorasProd Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos8 { get; set; }

        [Display(Name = "HorasProd Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos7 { get; set; }

        [Display(Name = "HorasProd Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos6 { get; set; }

        [Display(Name = "HorasProd Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos5 { get; set; }

        [Display(Name = "HorasProd Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos4 { get; set; }

        [Display(Name = "HorasProd Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos3 { get; set; }

        [Display(Name = "HorasProd Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAnoMenos2 { get; set; }

        [Display(Name = "HorasProd Ano")]       // MT
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float HorasProdAno { get; set; }

        [Display(Name = "RecNac Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos12 { get; set; }

        [Display(Name = "RecNac Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos11 { get; set; }

        [Display(Name = "RecNac Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos10 { get; set; }

        [Display(Name = "RecNac Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos9 { get; set; }

        [Display(Name = "RecNac Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos8 { get; set; }

        [Display(Name = "RecNac Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos7 { get; set; }

        [Display(Name = "RecNac Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos6 { get; set; }

        [Display(Name = "RecNac Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos5 { get; set; }

        [Display(Name = "RecNac Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos4 { get; set; }

        [Display(Name = "RecNac Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos3 { get; set; }

        [Display(Name = "RecNac Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAnoMenos2 { get; set; }

        [Display(Name = "RecNac Ano")]      // NF
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecNacAno { get; set; }

        [Display(Name = "RecExp Ano - 12")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos12 { get; set; }

        [Display(Name = "RecExp Ano - 11")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos11 { get; set; }

        [Display(Name = "RecExp Ano - 10")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos10 { get; set; }

        [Display(Name = "RecExp Ano - 9")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos9 { get; set; }

        [Display(Name = "RecExp Ano - 8")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos8 { get; set; }

        [Display(Name = "RecExp Ano - 7")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos7 { get; set; }

        [Display(Name = "RecExp Ano - 6")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos6 { get; set; }

        [Display(Name = "RecExp Ano - 5")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos5 { get; set; }

        [Display(Name = "RecExp Ano - 4")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos4 { get; set; }

        [Display(Name = "RecExp Ano - 3")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos3 { get; set; }

        [Display(Name = "RecExp Ano - 2")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAnoMenos2 { get; set; }

        [Display(Name = "RecExp Ano")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public float RecExpAno { get; set; }
    }
}