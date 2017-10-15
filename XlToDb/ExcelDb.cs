﻿using System;
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

            for (int k = 2; k < 4; k++)
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Produtos);
                Excel._Worksheet worksheet = workbook.Sheets[k];
                Excel.Range range = worksheet.UsedRange;

                try
                {
                    for (int i = 2; i <= range.Rows.Count + 1; i++)
                    {
                        string apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null
                            ? range.Cells[i, 1].Value2.ToString()
                            : "999999";
                        var teste = db.Produtos.Find(apelido);

                        if (teste == null)
                        {
                            var data = new Produto
                            {
                                Apelido = apelido,
                                Descricao = range.Cells[i, 2] != null && range.Cells[i, 2].Value2 != null
                                    ? range.Cells[i, 2].Value2.ToString()
                                    : "--",
                                UnidadeId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null
                                    ? Select.Unidade(range.Cells[i, 3].Value2.ToString())
                                    : 8,
                                TipoId = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null
                                    ? Select.Tipo(range.Cells[i, 4].Value2.ToString())
                                    : 4,
                                ClasseCustoId =
                                    range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null
                                        ? Select.ClasseCusto(range.Cells[i, 5].Value2.ToString())
                                        : 9,
                                CategoriaId =
                                    range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null
                                        ? Select.Categoria(range.Cells[i, 6].Value2.ToString())
                                        : 12,
                                FamiliaId = range.Cells[i, 7] != null && range.Cells[i, 7].Value2 != null
                                    ? Select.Familia(range.Cells[i, 7].Value2.ToString())
                                    : 15,
                                LinhaId = range.Cells[i, 8] != null && range.Cells[i, 8].Value2 != null
                                    ? Select.Linha(range.Cells[i, 8].Value2.ToString())
                                    : 12,
                                FlagProduto = k == 2 ? true : false
                            };

                            db.Produtos.Add(data);
                            db.SaveChanges();
                            Console.WriteLine($"{k}, {i}");
                        }

                        else DbLogger.Log(Reason.Info, $"Produto em duplicação: {apelido}");
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

            try
            {
                for (int i = 2; i < range.Rows.Count + 1; i++)
                {

                    var data = new Estrutura
                    {
                        Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "999999",
                        UnidadeId = range.Cells[i, 3] != null && range.Cells[i, 3].Value2 != null ? Select.Unidade(range.Cells[i, 3].Value2.ToString()) : 8,
                        QtdCusto = range.Cells[i, 4] != null && range.Cells[i, 4].Value2 != null ? (float)range.Cells[i, 4].Value2 : 0,
                        SequenciaId = range.Cells[i, 5] != null && range.Cells[i, 5].Value2 != null ? Select.Sequencia(range.Cells[i, 5].Value2.ToString()) : 9,
                        Item = range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null ? range.Cells[i, 6].Value2.ToString() : "999999",
                        Onera = range.Cells[i, 10] != null && range.Cells[i, 10].Value2 != null ? Select.Onera(range.Cells[i, 10].Value2.ToString()) : false,
                        Lote = range.Cells[i, 11] != null && range.Cells[i, 11].Value2 != null ? (float)range.Cells[i, 11].Value2 : 0,
                        Perda = range.Cells[i, 12] != null && range.Cells[i, 12].Value2 != null ? (float)range.Cells[i, 12].Value2 : 0,
                        Observacao = range.Cells[i, 13] != null && range.Cells[i, 13].Value2 != null ? range.Cells[i, 13].Value2.ToString() : "--",

                    };

                    db.Estruturas.Add(data);
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

            try
            {
                for (int i = 2; i < range.Rows.Count + 1; i++)
                {
                    int j = 0;
                    var data = new Cotacao
                    {
                        DateTime = DateTime.Now,
                        Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null
                            ? range.Cells[i, 1].Value2.ToString()
                            : "999999",
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

        public void Insumo()
        {
            var db = new EntityContext();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(Files.Insumos);
            Excel._Worksheet worksheet = workbook.Sheets[2];
            Excel.Range range = worksheet.UsedRange;

            try
            {
                for (int i = 2; i < range.Rows.Count + 1; i++)
                {
                    int j = 10;
                    var data = new Insumo
                    {
                        Apelido = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? range.Cells[i, 1].Value2.ToString() : "999999",
                        ProdutoId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? Select.Produto(range.Cells[i, 1].Value2.ToString()) : 4642,
                        Peso = range.Cells[i, 9] != null && range.Cells[i, 9].Value2 != null ? (float)range.Cells[i, 9].Value2 : 0,
                        CotacaoId = range.Cells[i, 1] != null && range.Cells[i, 1].Value2 != null ? Select.Cotacao(range.Cells[i, 1].Value2.ToString()) : 5032,
                        PrecoUsd = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        PrecoRs = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        Icms = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        Ipi = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        Pis = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        Cofins = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        DespExtra = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        DespImport = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        Ativo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Status(range.Cells[i, j].Value2.ToString()) : false,
                        FinalidadeId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Finalidade(range.Cells[i, j].Value2.ToString()) : 4,
                        UnddId = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? Select.Unidade(range.Cells[i, j].Value2.ToString()) : 8,
                        QtdUnddConsumo = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0,
                        QtdMltplCompra = range.Cells[i, ++j] != null && range.Cells[i, j].Value2 != null ? (float)range.Cells[i, j].Value2 : 0
                    };

                    db.Insumos.Add(data);
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
    }
}
