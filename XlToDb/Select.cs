using System;
using System.Linq;
using XlToDb.Model;

namespace XlToDb
{
    public static class Select
    {
        public static int Unidade(string celula)
        {
            var comp = celula.ToLower();

            if (comp == "cx" || comp == "caixa") return 1;
            if (comp == "pc" || comp == "pç" || comp == "peça" || comp == "peca") return 2;
            if (comp == "kg" || comp == "kilograma") return 3;
            if (comp == "m" || comp == "mt" || comp == "metro") return 4;
            if (comp == "rl" || comp == "rolo") return 5;
            if (comp == "ml" || comp == "milheiro") return 6;
            if (comp == "hr" || comp == "hora") return 9;
            if (comp == "un" || comp == "unidade") return 10;

            return 8;
        }

        public static int Tipo(string celula)
        {
            var comp = celula.Substring(0, 1).ToLower();

            if (comp == "a") return 1;
            if (comp == "b") return 2;

            return 4;
        }

        public static int ClasseCusto(string celula)
        {
            var comp = celula.Substring(0, 2);

            if (comp == "00") return 1;
            if (comp == "01") return 2;
            if (comp == "02") return 3;
            if (comp == "04") return 10;
            if (comp == "06") return 4;
            if (comp == "07") return 5;
            if (comp == "10") return 6;

            return 9;
        }

        public static int Categoria(string celula)
        {
            var db = new EntityContext();
            var comp = celula.Substring(0, 2);

            var result = db.Categorias.SingleOrDefault(c => c.Apelido == comp);
            if (result == null) return 12;
            return result.CategoriaId;
        }

        public static int Familia(string celula)
        {
            var comp = celula.Substring(0, 3);
            var db = new EntityContext();

            var result = db.Familias.SingleOrDefault(c => c.Apelido == comp);
            if (result == null) return 15;

            return result.FamiliaId;
        }

        public static int Linha(string celula)
        {
            var db = new EntityContext();
            var comp = celula.Substring(0, 4);

            var result = db.Linhas.SingleOrDefault(c => c.Apelido == comp);
            if (result == null) return 15;

            return result.LinhaId;
        }

        public static int GrupoRateio(string celula)
        {
            celula = celula.ToLower();

            if (celula == "fita") return 9;
            if (celula == "graxa") return 10;
            if (celula == "gxfpuro") return 11;
            if (celula == "revenda") return 12;
            if (celula == "tubo") return 13;
            if (celula == "sucata") return 14;
            if (celula == "descarte") return 15;
            if (celula == "gxfgraf") return 19;

            return 18;
        }

        public static int Sequencia(string celula)
        {
            var comp = celula.ToUpper();

            if (comp == "B") return 10;
            if (comp == "C") return 11;
            if (comp == "D") return 12;
            if (comp == "E1") return 13;
            if (comp == "E2") return 14;
            if (comp == "F") return 15;

            return 9;
        }

        public static bool Onera(string celula)
        {
            if (String.IsNullOrEmpty(celula)) return false;
            if (celula.ToLower() == "onera") return true;

            return false;
        }

        public static int Setor(string celula)
        {
            var db = new EntityContext();
            var setor = db.Setores.FirstOrDefault(s => s.Codigo == celula);
            if (setor == null) return 22;
            return setor.SetorId;
        }

        public static int Area(string celula)
        {
            var db = new EntityContext();
            var area = db.Areas.FirstOrDefault(a => a.Apelido == celula);
            if (area == null) return 1;
            return area.AreaId;
        }

        public static int MedidaFita(string largura, string comprimento)
        {
            var db = new EntityContext();
            int larg = (int)(float.Parse(largura) * 1000);
            int comp = int.Parse(comprimento);
            var medida = db.MedidaFitas.SingleOrDefault(m => m.LarguraMm == larg && m.ComprimentoMetros == comp);
            if (medida == null) return 1;
            return medida.MedidaFitaId;
        }

        public static bool Status(string celula)
        {
            if (celula.ToLower() == "ativo") return true;
            return false;
        }

        public static int Dominio(string celula)
        {
            var db = new EntityContext();
            var dominio = celula.Substring(2, celula.Length - 2).ToLower();
            var resposta = db.Dominios.SingleOrDefault(d => d.Descricao == dominio);
            if (resposta == null) return 1;
            return resposta.DominioId;
        }

        public static int TipoProducao(string celula)
        {
            int result = 3;

            if (celula == "IND") result = 1;
            else if (celula == "BEN") result = 2;

            return result;
        }

        public static int Pcp(string celula)
        {
            return celula == "PE" ? 1 : 2;
        }

        public static int Produto(string celula)
        {
            var db = new EntityContext();
            var resposta = db.Produtos.FirstOrDefault(p => p.Apelido == celula);
            if (resposta == null) return 16895;
            return resposta.Id;
        }

        //public static int Cotacao(string celula)
        //{
        //    var db = new EntityContext();
        //    var resposta = db.Cotacoes.SingleOrDefault(c => c.Apelido == celula);
        //    if (resposta == null) return 5032;
        //    return resposta.CotacaoId;
        //}

        public static int Finalidade(string celula)
        {
            var db = new EntityContext();
            var resposta = db.Finalidades.SingleOrDefault(c => c.Descricao == celula);
            if (resposta == null) return 4;
            return resposta.FinalidadeId;
        }

        public static int TipoAlteracao(string celula)
        {
            var db = new EntityContext();
            var resposta = db.TiposAlteracao.SingleOrDefault(c => c.Descricao == celula);
            if (resposta == null) return 4;
            return resposta.TipoAlteracaoId;
        }

        public static int Embalagem(string celula)
        {
            var db = new EntityContext();
            var embalagem = db.Embals.SingleOrDefault(e => e.Sigla == celula);
            int result = embalagem == null ? 16 : embalagem.Id;

            return result;
        }

        public static int Resina(string celula)
        {
            var db = new EntityContext();
            var resina = db.Resinas.SingleOrDefault(r => r.Descricao == celula);
            int result = resina == null ? 3 : resina.Id;

            return result;
        }

        public static int MedidaFita(string celula)
        {
            var db = new EntityContext();
            var medidaFita = db.MedidaFitas.SingleOrDefault(mf => mf.Apelido == celula);
            var result = medidaFita == null
                ? 32
                : medidaFita.MedidaFitaId;

            return result;
        }

        public static int Prensa(string celula)
        {
            var db = new EntityContext();
            var prensa = db.PrensasPreForma.SingleOrDefault(p => p.Apelido == celula);
            var result = prensa == null
                ? 4
                : prensa.Id;

            return result;
        }

        public static int Extrusora(string celula)
        {
            var db = new EntityContext();
            var extrusora = db.Extrusoras.SingleOrDefault(p => p.Apelido == celula);
            var result = extrusora == null
                ? 3
                : extrusora.Id;

            return result;
        }

        public static int Fabricante(string celula)
        {
            var db = new EntityContext();
            var fabricante = db.Fabricantes.SingleOrDefault(p => p.Apelido == celula);
            var result = fabricante == null
                ? 5
                : fabricante.Id;

            return result;
        }

        public static int Insumo(string celula)
        {
            var db = new EntityContext();
            var insumo = db.Insumos.SingleOrDefault(p => p.Apelido == celula);
            var result = insumo == null
                ? 2450
                : insumo.InsumoId;

            return result;
        }

        public static int ResinaBase(string celula)
        {
            var db = new EntityContext();
            var resina = db.ResinasBase.SingleOrDefault(p => p.Apelido == celula);
            var result = resina == null
                ? 4
                : resina.Id;

            return result;
        }

        public static int Serie(string celula)
        {
            var db = new EntityContext();
            var serie = db.Series.SingleOrDefault(p => p.Apelido == celula);
            var result = serie == null
                ? 17
                : serie.Id;

            return result;
        }

        public static int Carga(string celula)
        {
            celula = celula.Substring(0, 3);
            var db = new EntityContext();
            var carga = db.Cargas.SingleOrDefault(p => p.Apelido == celula);
            var result = carga == null
                ? 15
                : carga.Id;

            return result;
        }

        public static bool Sinterizado(string celula)
        {
            bool result = celula == "SINT"
                ? true
                : false;

            return result;
        }

        public static bool ProcessoContinuo(string celula)
        {
            bool result = celula == "1-Continuo"
                ? true
                : false;

            return result;
        }

        public static bool TesteEstqEstoq(string celula)
        {
            bool result = celula == "sim"
                ? true
                : false;

            return result;
        }

        public static int CondicaoPreco(string celula)
        {
            int result = 4;
            var db = new EntityContext();
            var condicao = db.CondicoesPrecos.SingleOrDefault(c => c.Apelido == celula);
            if (condicao != null) result = condicao.Id;

            return result;
        }

        public static bool DespesaPadrao(string celula)
        {
            bool result = celula == "sim"
                ? true
                : false;

            return result;
        }

        public static int Operacao(string celula)
        {
            int result = 583;
            var db = new EntityContext();
            var operacao = db.Operacoes.SingleOrDefault(o => o.CodigoOperacao == celula);
            if (operacao != null) result = operacao.OperacaoId;

            return result;
        }
    }
}
