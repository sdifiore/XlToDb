using System;

namespace XlToDb
{
    class Program
    {
        static void Main(string[] args)
        {
            var xl = new ExcelDb();
            //xl.Produto();
            //xl.Estrutura();
            //xl.Operacao();
            //xl.Folha();
            //xl.QtdEmbalagem();
            //xl.ParteProduto();
            //xl.Cotacao();
            //xl.Insumo();
            //xl.Alteracao();
            xl.UpdateTipo();
        }
    }
}
