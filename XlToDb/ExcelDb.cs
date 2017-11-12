using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using XlToDb.Model;

namespace XlToDb
{
    public class ExcelDb
    {
        public void Produto()
        {
            var db = new EntityContext();


            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Produtos);
            Excel._Worksheet worksheet = workbook.Sheets[3];
            Excel.Range range = worksheet.UsedRange;

                for (int i = 2; i <= range.Rows.Count + 1; i++)
                {
                    var data = new Produto();

                    int j = 1;

                    data.Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null
                        ? range.Cells[i, 1].Value2.ToString()
                        : "999999";
                    data.Descricao = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? range.Cells[i, j].Value2.ToString()
                        : "--";
                    data.UnidadeId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Unidade(range.Cells[i, j].Value2.ToString())
                        : 8;
                    data.TipoId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Tipo(range.Cells[i, j].Value2.ToString())
                        : 4;
                    data.ClasseCustoId =
                        range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                            ? Select.ClasseCusto(range.Cells[i, j].Value2.ToString())
                            : 9;
                    data.CategoriaId =
                        range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                            ? Select.Categoria(range.Cells[i, j].Value2.ToString())
                            : 12;
                    data.FamiliaId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Familia(range.Cells[i, j].Value2.ToString())
                        : 15;
                    data.LinhaId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Linha(range.Cells[i, j].Value2.ToString())
                        : 12;
                    data.GrupoRateioId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.GrupoRateio(range.Cells[i, j].Value2.ToString())
                        : 18;
                    data.PesoLiquido = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? (float) range.Cells[i, j].Value2
                        : 0;
                    data.Ativo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Status(range.Cells[i, j].Value2.ToString())
                        : false;
                    data.Ipi = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? (float) range.Cells[i, j].Value2
                        : 0;
                    data.QtdUnid = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? (int) range.Cells[i, j].Value2
                        : 0;
                    data.DominioId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Dominio(range.Cells[i, j].Value2.ToString())
                        : 1;
                    data.TipoProdId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.TipoProducao(range.Cells[i, j].Value2.ToString())
                        : 1;
                    data.PcpId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Pcp(range.Cells[i, j].Value2.ToString())
                        : 1;
                    data.QtUnPorUnArmz = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0 ;
                    data.MedidaFitaId = 32;

                    db.Produtos.Add(data);
                    db.SaveChanges();
                    Console.WriteLine(i);
            }
             
        }

        public void ParteProduto()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.ParteProduto);
            Excel._Worksheet worksheet = workbook.Sheets[3];
            Excel.Range range = worksheet.UsedRange;

            try
            {
                for (int i = 2; i < range.Rows.Count + 1; i++)
                {
                    int j = 8;
                    var data = new ParteProduto
                    {
                        GrupoRateioId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.GrupoRateio(range.Cells[i, j].Value2.ToString()) : 18,
                        Peso = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        Status = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Status(range.Cells[i, j].Value2.ToString()) : false,
                        Ipi = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        QtdUndd = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                        DominioId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Dominio(range.Cells[i, j].Value2.ToString()) : 1,
                        TipoProducaoId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.TipoProducao(range.Cells[i, j].Value2.ToString()) : 1,
                        PcpId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Pcp(range.Cells[i, j].Value2.ToString()) : 1,
                        QtdUndArmz = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0,
                        ProdutoId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? Select.Produto(range.Cells[i, 1].Value2.ToString()) : 4642
                    };

                    db.ParteProdutos.Add(data);
                    db.SaveChanges();
                    Console.WriteLine(i);
                }
            }
            catch (Exception ex)
            {
                DbLogger.Log(Reason.Error, ex.Message);
            }
            finally
            {
                xlApp.Quit();
                workbook = null;
                worksheet = null;
                range = null;
            }
        }

        public void Estrutura()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Estrutura);
            Excel._Worksheet worksheet = workbook.Sheets[4];
            Excel.Range range = worksheet.UsedRange;

            
                for (int i = 2; i < range.Rows.Count + 1; i++)
                {

                    var data = new Estrutura();

                    data.ProdutoId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null
                        ? Select.Produto(range.Cells[i, 1].Value2.ToString())
                        : 14844;
                    data.UnidadeId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null
                        ? Select.Unidade(range.Cells[i, 3].Value2.ToString())
                        : 8;
                    data.QtdCusto = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null
                        ? (float) range.Cells[i, 4].Value2
                        : 0;
                    data.SequenciaId = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null
                        ? Select.Sequencia(range.Cells[i, 5].Value2.ToString())
                        : 9;
                    data.Item = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null
                        ? range.Cells[i, 6].Value2.ToString()
                        : "999999";
                    data.Onera = range.Cells[i, 10] != null && range.Cells[i, 10].Value2 != null
                        ? Select.Onera(range.Cells[i, 10].Value2.ToString())
                        : false;
                    data.Lote = range.Cells[i, 11] != null && range.Cells[i, 11].Value2 != null
                        ? (float) range.Cells[i, 11].Value2
                        : 0;
                    data.Perda = range.Cells[i, 12] != null && range.Cells[i, 12].Value2 != null
                        ? (float) range.Cells[i, 12].Value2
                        : 0;
                    data.Observacao = range.Cells[i, 13] != null && range.Cells[i, 13].Value2 != null
                        ? range.Cells[i, 13].Value2.ToString()
                        : "--";

                    db.Estruturas.Add(data);
                    db.SaveChanges();
                    Console.WriteLine(i);
                }
            
            
        }

        public void QtdLoteComp()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Estrutura);
            Excel._Worksheet worksheet = workbook.Sheets[4];
            Excel.Range range = worksheet.UsedRange;
            int i = 1;
            var db = new EntityContext();
            var model = db.Estruturas;

            foreach (var register in model)
            {
                register.Lote = range.Cells[++i, 11] != null && range.Cells[i, 11].Value2 != null
                    ? (float) range.Cells[i, 11].Value2
                    : 0;
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void Perdas()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Estrutura);
            Excel._Worksheet worksheet = workbook.Sheets[4];
            Excel.Range range = worksheet.UsedRange;
            int i = 1;
            var db = new EntityContext();
            var model = db.Estruturas;

            foreach (var register in model)
            {
                register.Perda = range.Cells[++i, 12] != null && range.Cells[i, 12].Value2 != null
                    ? (float)range.Cells[i, 12].Value2
                    : 0;
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

            try
            {
                for (int i = 2; i < 53; i++)
                {
                    var operacao = new Operacao
                    {
                        CodigoOperacao = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "999999",
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
            catch (Exception ex)
            {
                DbLogger.Log(Reason.Error, ex.Message);
            }
            finally
            {
                xlApp.Quit();
                workbook = null;
                worksheet = null;
                range = null;
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

            try
            {
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
            catch (Exception ex)
            {
                DbLogger.Log(Reason.Error, ex.Message);
            }
            finally
            {
                xlApp.Quit();
                workbook = null;
                worksheet = null;
                range = null;
            }
        }

        public void QtdEmbalagem()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.QtdEmbalagem);
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            try
            {
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
            catch (Exception ex)
            {
                DbLogger.Log(Reason.Error, ex.Message);
            }
            finally
            {
                xlApp.Quit();
                workbook = null;
                worksheet = null;
                range = null;
            }
        }

        public void Cotacao()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Cotacao);
            Excel._Worksheet worksheet = workbook.Sheets[2];
            Excel.Range range = worksheet.UsedRange;

            {
                for (int i = 2; i < range.Rows.Count + 1; i++)
                {
                    int j = 0;
                    var data = new Cotacao
                    {
                        DateTime = DateTime.Now,
                        Descricao = range.Cells[i, 10] != null && range.Cells[i, 10].Value2 != null
                            ? range.Cells[i, 10].Value2.ToString()
                            : "--",
                        ProdutoId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null
                            ? Select.Produto(range.Cells[i, 1].Value2.ToString())
                            : 4642
                    };

                    db.Cotacoes.Add(data);
                    db.SaveChanges();
                    Console.WriteLine(i);
                }
            }
        }

        public void Insumo()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Insumos);
            Excel._Worksheet worksheet = workbook.Sheets[2];
            Excel.Range range = worksheet.UsedRange;


                for (int i = 2; i < range.Rows.Count + 1; i++)
                {
                    int j = 0;
                    var data = new Insumo();

                   data.Apelido = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? range.Cells[i, j].Value2.ToString()
                        : "999999";
                    data.Descricao = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? range.Cells[i, j].Value2.ToString()
                        : "--";
                    data.UnidadeId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Unidade(range.Cells[i, j].Value2.ToString())
                        : 8;
                    data.TipoId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Tipo(range.Cells[i, j].Value2.ToString())
                        : 4;
                    data.ClasseCustoId =
                        range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                            ? Select.ClasseCusto(range.Cells[i, j].Value2.ToString())
                            : 9;
                    data.CategoriaId =
                        range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                            ? Select.Categoria(range.Cells[i, j].Value2.ToString())
                            : 12;
                    data.FamiliaId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Familia(range.Cells[i, j].Value2.ToString())
                        : 15;
                    data.LinhaId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Linha(range.Cells[i, j].Value2.ToString())
                        : 12;
                    data.Peso = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? (float) range.Cells[i, j].Value2
                        : 0;
                    data.Cotacao = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? range.Cells[i, j].Value2.ToString()
                        : "--";
                    data.PrecoUsd = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.PrecoRs = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.Icms = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.Ipi = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.Pis = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.Cofins = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.DespExtra = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.DespImport = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.Ativo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null
                        ? Select.Status(range.Cells[i, j].Value2.ToString())
                        : false;
                    data.FinalidadeId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Finalidade(range.Cells[i, j].Value2.ToString()) : 4;
                    data.UnddId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Unidade(range.Cells[i, j].Value2.ToString()) : 8;
                    data.QtdUnddConsumo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.QtdMltplCompra = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.FormaPgto = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? range.Cells[i, j].Value2.ToString() : "--";
                    data.Prazo1 = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                    data.Prazo2 = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                    data.PctPgto1 = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                    data.ImportPzPagDesp = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;

                    db.Insumos.Add(data);
                    db.SaveChanges();
                    Console.WriteLine(i);
                }
        }

        public void Alteracao()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Alteracao);
            Excel._Worksheet worksheet = workbook.Sheets[7];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < 122; i++)
            {
                var data = new Ajuste
                {
                    OrigemId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? Select.Produto(range.Cells[i, 1].Value2.ToString()) : 4642,
                    UnidadeDeId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null ? Select.Unidade(range.Cells[i, 3].Value2.ToString()) : 8,
                    AtualId = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null ? Select.Produto(range.Cells[i, 4].Value2.ToString()) : 4642,
                    UnidadeParaId = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null ? Select.Unidade(range.Cells[i, 6].Value2.ToString()) : 8,
                    Fator = range.Cells[i, 7] != null && range.Cells[i, 7].Value2 != null ? (float)range.Cells[i, 7].Value2 : 0,
                    TipoAlteracaoId = range.Cells[i, 8] != null && range.Cells[i, 8].Value2 != null ? Select.TipoAlteracao(range.Cells[i, 8].Value2.ToString()) : 4,
                    Medida = range.Cells[i, 9] != null && range.Cells[i, 9].Value2 != null ? (float)range.Cells[i, 9].Value2 : 0
                };

                db.Ajustes.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void UpdateTipo()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Produtos);
            Excel._Worksheet worksheet = workbook.Sheets[3];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < range.Count + 1; i++)
            {
                string comp = range.Cells[i, 1].Value2.ToString();
                var data = db.Produtos.Single(p => p.Apelido == comp);
                data.TipoId = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null
                    ? Select.Tipo(range.Cells[i, 4].Value2.ToString())
                    : 4;
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void EncapTubos()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.EncapTubos);
            Excel._Worksheet worksheet = workbook.Sheets[12];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < 9; i++)
            {
                int j = 2;
                var data = new EncapTubo();

                data.ProdutoId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? Select.Produto(range.Cells[i, 1].Value2.ToString()) : 4642;
                data.UnidadeId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Unidade(range.Cells[i, j].Value2.ToString()) : 8;
                data.DextRevest = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                data.DintRevest = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                data.ResinaBase = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? range.Cells[i, j].Value2.ToString() : "--";
                data.Aditivo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? range.Cells[i, j].Value2.ToString() : "--";
                data.DenRev = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                data.PesoRevest = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                data.VelRevest = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.PctCarga = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                

                db.EncapTubos.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void Graxas()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Graxas);
            Excel._Worksheet worksheet = workbook.Sheets[11];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 3; i < 31; i++)
            {
                var data = new Graxa();

                data.Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "999999";
                data.Descricao = range.Cells[i, 2] != null && range.Cells[i, 2].Value2 != null ? range.Cells[i, 2].Value2.ToString() : "--";
                data.EmbalagemId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null ? Select.Embalagem(range.Cells[i, 3].Value2.ToString()) : 9;
                data.Peso = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null ? (float)range.Cells[i, 4].Value2 : 0;
                data.PctSilicone = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null ? (float)range.Cells[i, 5].Value2 : 0;
                data.PctSilica = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null ? (float)range.Cells[i, 6].Value2 : 0;
                data.ResinaId = range.Cells[i, 8] != null && range.Cells[i, 8].Value2 != null ? Select.Resina(range.Cells[i, 8].Value2.ToString()) : 3;
                data.EmbalagemMedida = range.Cells[i, 11] != null && range.Cells[i, 11].Value2 != null ? (float)range.Cells[i, 11].Value2 : 0;
                data.Rotulagem = range.Cells[i, 12] != null && range.Cells[i, 12].Value2 != null ? (float)range.Cells[i, 12].Value2 : 0;

                db.Graxas.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void AjusteProdutos()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.AjusteProduto);
            Excel._Worksheet worksheet = workbook.Sheets[6];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < range.Rows.Count + 1; i++)
            {
                int id = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? Select.Produto(range.Cells[i, 1].Value2.ToString()) : 14844;
                var data = db.Produtos.SingleOrDefault(p => p.Id == id);
                if (data != null)
                {
                    data.MedidaFitaId = range.Cells[i, 31] != null && range.Cells[i, 31].Value2 != null ? Select.MedidaFita(range.Cells[i, 31].Value2.ToString()) : 32;
                }

                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void PadraoFixo()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.PadraoFixo);
            Excel._Worksheet worksheet = workbook.Sheets[3];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 3; i < 38; i++)
            {
                var data = new PadraoFixo
                {
                    Descricao = range.Cells[i, 17] != null && range.Cells[i, 17].Value2 != null ? range.Cells[i, 17].Value2.ToString() : "--",
                    Valor = range.Cells[i, 18] != null && range.Cells[i, 18].Value2 != null ? (float)range.Cells[i, 18].Value2 : 0
                };

                db.PadroesFixos.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void PreForma()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.PreForma);
            Excel._Worksheet worksheet = workbook.Sheets[3];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 3; i < 11; i++)
            {
                int j = 0;
                var data = new PreForma();
                data.PreFormaNum = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.FormaDiamE = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.VaretaDiamI = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.Medidas = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? range.Cells[i, j].Value2.ToString() : "--";
                data.Comprimento = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.Tup = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.PrensaPreFormaId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Prensa(range.Cells[i, j].Value2.ToString()) : 4;
                data.Preparo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.TrocaPf = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0;
                data.ExtrusoraId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Extrusora(range.Cells[i, j].Value2.ToString()) : 3;
                data.DiamPistaoHidraulico = range.Cells[i, 14] != null && range.Cells[i, 14].Value2 != null ? (float)range.Cells[i, 14].Value2 : 0;

                db.PreFormas.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }

        public void ResinaPtfe()
        {
            var db = new EntityContext();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Resina);
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            for (int i = 2; i < range.Rows.Count + 1; i++)
            {
                int j = 0;
                var data = new ResinaPtfe();
                data.Ref = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (int)range.Cells[i, j].Value2 : 0;
                data.Referencia = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? range.Cells[i, j].Value2.ToString() : "--";
                data.FabricanteId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Fabricante(range.Cells[i, j].Value2.ToString()) : 5;
                data.ResinaBaseId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.ResinaBase(range.Cells[i, j].Value2.ToString()) : 4;
                data.InsumoId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Insumo(range.Cells[i, j].Value2.ToString()) : 2450;
                data.MaxRr = range.Cells[i, 12] != null && range.Cells[i, 12].Value2 != null ? (int)range.Cells[i, 12].Value2 : 0;
                data.Classificacao = range.Cells[i, 13] != null && range.Cells[i, 13].Value2 != null ? range.Cells[i, 13].Value2.ToString() : "--";
                data.MaxRrAntiga = range.Cells[i, 14] != null && range.Cells[i, 14].Value2 != null ? (int)range.Cells[i, 14].Value2 : 0;

                db.ResinasPtfe.Add(data);
                db.SaveChanges();
                Console.WriteLine(i);
            }
        }
    }
}
