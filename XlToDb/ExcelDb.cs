using System;
using System.CodeDom;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using XlToDb.Model;

namespace XlToDb
{
    public class ExcelDb
    {
        public void Insumo()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Insumos);
            Excel._Worksheet worksheet = workbook.Sheets[2];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < range.Rows.Count; i++)
            {
                var data = new Insumo
                {
                    Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "0000000000",
                    Descricao = range.Cells[i, 2] != null && range.Cells[i, 2].Value2 != null ? range.Cells[i, 2].Value2.ToString() : "--",
                    UnidadeId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null ? Select.Unidade(range.Cells[i, 3].Value2.ToString()) : 8,
                    TipoId = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null ? Select.Tipo(range.Cells[i, 4].Value2.ToString()) : 4,
                    ClasseCustoId = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null ? Select.ClasseCusto(range.Cells[i, 5].Value2.ToString()) : 9,
                    CategoriaId = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null ? Select.Categoria(range.Cells[i, 6].Value2.ToString()) : 12,
                    FamiliaId = range.Cells[i, 7] != null && range.Cells[i, 7].Value2 != null ? Select.Familia(range.Cells[i, 7].Value2.ToString()) : 15,
                    LinhaId = range.Cells[i, 8] != null && range.Cells[i, 8].Value2 != null ? Select.Linha(range.Cells[i, 8].Value2.ToString()) : 12,
                    Peso = range.Cells[i, 9] != null && range.Cells[i, 9].Value2 != null ? (float)range.Cells[i, 9].Value2 : 0,
                };
            }
        }

        public void Produto()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Produtos);
            Excel._Worksheet worksheet = workbook.Sheets[3];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i <= range.Rows.Count; i++)
            {

                var data = new Produto
                {
                    Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "0000000000",
                    Descricao = range.Cells[i, 2] != null && range.Cells[i, 2].Value2 != null ? range.Cells[i, 2].Value2.ToString() : "--",
                    UnidadeId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null ? Select.Unidade(range.Cells[i, 3].Value2.ToString()) : 8,
                    TipoId = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null ? Select.Tipo(range.Cells[i, 4].Value2.ToString()) : 4,
                    ClasseCustoId = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null ? Select.ClasseCusto(range.Cells[i, 5].Value2.ToString()) : 9,
                    CategoriaId = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null ? Select.Categoria(range.Cells[i, 6].Value2.ToString()) : 12,
                    FamiliaId = range.Cells[i, 7] != null && range.Cells[i, 7].Value2 != null ? Select.Familia(range.Cells[i, 7].Value2.ToString()) : 15,
                    LinhaId = range.Cells[i, 8] != null && range.Cells[i, 8].Value2 != null ? Select.Linha(range.Cells[i, 8].Value2.ToString()) : 12,
                    GrupoRateioId = range.Cells[i, 9] != null && range.Cells[i, 9].Value2 != null ? Select.GrupoRateio(range.Cells[i, 9].Value2.ToString(), 3) : 18
                };


                db.Produtos.Add(data);
                db.SaveChanges();
            } 
        }

        public void Estrutura()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Estrutura);
            Excel._Worksheet worksheet = workbook.Sheets[4];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < range.Rows.Count; i++)
            {
                int produtoId;
                Produto produto;

                if (range.Cells[i, 6] == null || range.Cells[i, 6].Value2 == null) produtoId = 1637;
                else
                {
                    string celula = range.Cells[i, 6].Value2.ToString();
                    produto = db.Produtos.FirstOrDefault(p => p.Apelido == celula);
                    if (produto == null) produtoId = 1637;
                    else produtoId = produto.Id;
                }

                    var data = new Estrutura
                {
                    Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "0000000000",
                    ProdutoId = produtoId,
                    Onera = range.Cells[i, 10] != null && range.Cells[i, 10].Value2 != null ? Select.Onera(range.Cells[i, 10].Value2.ToString()) : false,
                    Observacao = range.Cells[i, 13] != null && range.Cells[i, 13].Value2 != null ? range.Cells[i, 13].Value2.ToString() : "--",
                    SequenciaId = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null ? Select.Sequencia(range.Cells[i, 5].Value2.ToString()) : 9
                    };

                Console.WriteLine(i);
                ;
                db.Estruturas.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void Operacao()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Operacao);
            Excel._Worksheet worksheet = workbook.Sheets[5];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < 53; i++)
            {
                var operacao = new Operacao
                {
                    CodigoOperacao = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "0000000000",
                    SetorProducao = range.Cells[i, 2] != null && range.Cells[i, 2].Value2 != null ? range.Cells[i, 2].Value2.ToString() : "--",
                    Descricao = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null ? range.Cells[i, 3].Value2.ToString() : "--",
                    TaxaOcupacao = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null ? (float)range.Cells[i, 4].Value2 : 0,
                    Comentario = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null ? range.Cells[i, 5].Value2.ToString() : "--",
                    QtdMaquinas = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null ? (int)range.Cells[i, 6].Value2 : 1,
                    Custo = range.Cells[i, 7] != null && range.Cells[i, 7].Value2 != null ? (float)range.Cells[i, 7].Value2 : 0,
                    SetorId = range.Cells[i, 8] != null && range.Cells[i, 8].Value2 != null ? Select.Setor(range.Cells[i, 8].Value2.ToString(), i) : 22
                };

                db.Operacoes.Add(operacao);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void Folha()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Folha);
            Excel._Worksheet worksheet = workbook.Sheets[9];
            Excel.Range range = worksheet.UsedRange;

            var stack = new Stack<int>();
            stack.Push(15);
            stack.Push(14);
            stack.Push(13);
            stack.Push(12);
            stack.Push(11);
            stack.Push(9);
            stack.Push(8);
            stack.Push(7);
            stack.Push(5);
            stack.Push(4);
            stack.Push(3);
            stack.Push(2);

            for (int i = 1; i < 13; i++)
            {
                int j = 21;
                int k = (int)stack.Pop();
                var data = new CustoFolha
                {
                    Data = DateTime.Now,
                    Salario = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    Ferias = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    DecimoTerceiro = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    Plr = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    Fgts = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    Inss = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    DespAgencia = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    ConvMedico = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    VAlimentacao = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    VTransporte = range.Cells[++j, k] != null && range.Cells[j, k].Value2 != null ? (float)range.Cells[j, k].Value2 : 0,
                    AreaId = range.Cells[19, k] != null && range.Cells[19, k].Value2 != null ? Select.Area(range.Cells[19, k].Value2.ToString()) : 22
                };

                db.CustoFolhas.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void QtdEmbalagem()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.QtdEmbalagem);
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 3; i < 27; i++)
            {
                int j = 13;

                var data = new QtdEmbalagem
                {
                    CartuchoRolCx = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                    CartuchoCxPlt = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                    DisplayRolCx = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                    CarretelRolCx = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                    CarretelCxPlt = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                    MedidaFitasId = Select.MedidaFita(range.Cells[i, ++j].Value2.ToString(), range.Cells[i, ++j].Value2.ToString())
                };

                db.QtdEmbalagems.Add(data);
                db.SaveChanges();

            }
        }


    }
}
