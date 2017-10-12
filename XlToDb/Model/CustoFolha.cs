using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace XlToDb.Model
{
    public class CustoFolha
    {
        public int CustoFolhaId { get; set; }

        [Display(Name = "Data")]
        [DataType(DataType.Date)]
        public DateTime Data { get; set; }

        [Display(Name = "Salários brutos")]
        public float Salario { get; set; }

        [Display(Name = "Férias")]
        public float Ferias { get; set; }

        [Display(Name = "13º")]
        public float DecimoTerceiro { get; set; }

        [Display(Name = "PLR")]
        public float Plr { get; set; }

        [Display(Name = "FGTS")]
        public float Fgts { get; set; }

        [Display(Name = "INSS")]
        public float Inss { get; set; }

        [Display(Name = "Desp. Agência")]
        public float DespAgencia { get; set; }

        [Display(Name = "Conv. Médico")]
        public float ConvMedico { get; set; }

        [Display(Name = "Vale Alimentação")]
        public float VAlimentacao { get; set; }

        [Display(Name = "Vale Transporte")]
        public float VTransporte { get; set; }

        [Display(Name = "Área")]
        public int AreaId { get; set; }

        public Area Area { get; set; }
    }
}
