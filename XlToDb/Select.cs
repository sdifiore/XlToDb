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

            return 8;
        }

        public static int Tipo(string celula)
        {
            var comp = celula.Substring(0, 2).ToLower();

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
            if (comp == "06") return 4;
            if (comp == "07") return 5;
            if (comp == "10") return 6;

            return 9;
        }

        public static int Categoria(string celula)
        {
            var comp = celula.Substring(0, 2);

            if (comp == "20") return 1;
            if (comp == "50") return 2;
            if (comp == "51") return 3;
            if (comp == "52") return 4;
            if (comp == "60") return 5;
            if (comp == "61") return 6;
            if (comp == "71") return 7;
            if (comp == "82") return 8;
            if (comp == "91") return 9;

            return 12;
        }

        public static int Familia(string celula)
        {
            var comp = celula.Substring(0, 3);

            if (comp == "201") return 1;
            if (comp == "501") return 2;
            if (comp == "502") return 2;
            if (comp == "503") return 4;
            if (comp == "511") return 5;
            if (comp == "512") return 6;
            if (comp == "601") return 7;
            if (comp == "602") return 8;
            if (comp == "606") return 9;
            if (comp == "607") return 10;
            if (comp == "610") return 11;
            if (comp == "611") return 12;
            if (comp == "613") return 13;

            return 15;
        }

        public static int Linha(string celula)
        {
            var comp = celula.Substring(0, 4);

            if (comp == "1015") return 1;
            if (comp == "1021") return 2;
            if (comp == "1024") return 3;
            if (comp == "1025") return 4;
            if (comp == "2112") return 5;
            if (comp == "2113") return 6;
            if (comp == "2114") return 7;
            if (comp == "2115") return 8;
            if (comp == "5011") return 9;

            return 12;
        }

        public static int GrupoRateio(string celula)
        {
            if (celula == "fita") return 9;
            if (celula == "graxa") return 10;
            if (celula == "gxfpuro") return 11;
            if (celula == "revenda") return 12;
            if (celula == "tubo") return 13;
            if (celula == "sucata") return 14;
            if (celula == "descarte") return 15;

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

        public static int Setor(string celula, int i)
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
            return celula == "IND" ? 1 : 2;
        }

        public static int Pcp(string celula)
        {
            return celula == "PE" ? 1 : 2;
        }

        public static int Produto(string celula)
        {
            var db = new EntityContext();
            var resposta = db.Produtos.SingleOrDefault(p => p.Apelido == celula);
            if (resposta == null) return 4642;
            return resposta.Id;
        }

        public static int Cotacao(string celula)
        {
            var db = new EntityContext();
            var resposta = db.Cotacoes.SingleOrDefault(c => c.Apelido == celula);
            if (resposta == null) return 5032;
            return resposta.CotacaoId;
        }

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
    }
}
