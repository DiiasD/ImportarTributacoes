using ExcelDataReader;
using FirebirdSql.Data.FirebirdClient;
using OfficeOpenXml;
using System.Globalization;
using System.Text;

namespace ImportarTributacoes
{
    public class ImportacaoService : IDisposable
    {
        private readonly string _caminhoBanco;
        private readonly string _caminhoPlanilha;
        private readonly int _linhaInicial;

        public ImportacaoService(string caminhoBanco, string caminhoPlanilha, int linhaInicial)
        {
            _caminhoBanco = caminhoBanco ?? throw new ArgumentNullException(nameof(caminhoBanco));
            _caminhoPlanilha = caminhoPlanilha ?? throw new ArgumentNullException(nameof(caminhoPlanilha));
            _linhaInicial = linhaInicial >= 1 ? linhaInicial : throw new ArgumentException("Linha inicial deve ser >= 1");
        }

        public void Executar()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var produtos = LerPlanilha(_caminhoPlanilha);
            AtualizarBanco(produtos);
        }

        private List<TributacaoImportacao> LerPlanilha(string caminhoPlanilha)
        {
            var lista = new List<TributacaoImportacao>();

            using var package = new ExcelPackage(new FileInfo(caminhoPlanilha));
            var ws = package.Workbook.Worksheets[0]
                     ?? throw new InvalidOperationException("A planilha não contém nenhuma aba.");

            // Cabeçalho fixo na linha 3, dados a partir da linha configurada pelo usuário
            Console.WriteLine("Iniciando mapeamento de colunas...");
            var colunas = MapearCabecalhos(ws, 3);

            int ultimaLinha = ws.Dimension?.End.Row ?? 3;
            for (int row = _linhaInicial; row <= ultimaLinha; row++)
            {
                var codigoProduto = ws.Cells[row, colunas["PRODUTO"]].GetValue<string>()?.Trim();

                if (string.IsNullOrWhiteSpace(codigoProduto) &&
                    string.IsNullOrWhiteSpace(ws.Cells[row, colunas["NCM"]].GetValue<string>()?.Trim()))
                    continue; // Ignora linhas completamente vazias

                var item = new TributacaoImportacao
                {
                    Produto = codigoProduto ?? string.Empty,
                    Ncm = ws.Cells[row, colunas["NCM"]].GetValue<string>()?.Trim() ?? string.Empty,
                    Cst = ws.Cells[row, colunas["CST"]].GetValue<string>()?.Trim() ?? string.Empty,
                    ClassTrib = ws.Cells[row, colunas["CLASS_TRIB"]].GetValue<string>()?.Trim() ?? string.Empty,
                    AliqIbsUf = LerDecimal(ws, row, colunas["IBS_UF"]),
                    AliqIbsMun = LerDecimal(ws, row, colunas["IBS_MUN"]),
                    AliqCbs = LerDecimal(ws, row, colunas["CBS"])
                };

                lista.Add(item);
            }

            return lista;
        }

        private decimal LerDecimal(ExcelWorksheet ws, int linha, int coluna)
        {
            var valor = ws.Cells[linha, coluna].Value;

            if (valor == null) return 0m;

            if (decimal.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.GetCultureInfo("pt-BR"), out var resultado))
                return resultado;

            // Tenta com InvariantCulture (caso tenha ponto como separador)
            if (decimal.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out resultado))
                return resultado;

            return 0m;
        }

        private Dictionary<string, int> MapearCabecalhos(ExcelWorksheet ws, int linhaCabecalho)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            int ultimaColuna = ws.Dimension?.End.Column ?? 0;
            for (int col = 1; col <= ultimaColuna; col++)
            {
                var header = ws.Cells[linhaCabecalho, col].GetValue<string>()?.Trim().ToUpperInvariant();
                if (!string.IsNullOrEmpty(header))
                {
                    // Mapeia variações comuns dos cabeçalhos
                    string chave = header switch
                    {
                        "PRODUTO" or "CÓDIGO" or "CODIGO" => "PRODUTO",
                        "NCM" or "NCM/SH" or "NCMS" => "NCM",
                        "CST" or "CST_IBS" or "CST IBS/CBS" => "CST",
                        "CLASS_TRIB" or "CLASS. TRIB." or "CLASSTRIB" or "CLASSIFICAÇÃO TRIBUTÁRIA" => "CLASS_TRIB",
                        "IBS_UF" or "ALÍQ. IBS UF" or "IBS UF" or "ALÍQ IBS UF" => "IBS_UF",
                        "IBS_MUN" or "ALÍQ. IBS MUN." or "IBS MUN" or "ALÍQ IBS MUN" => "IBS_MUN",
                        "CBS" or "ALÍQ. CBS" or "ALÍQ CBS" => "CBS",
                        _ => header
                    };

                    if (!map.ContainsKey(chave))
                        map[chave] = col;
                }
            }

            // Validação das colunas obrigatórias
            var obrigatorias = new[] { "PRODUTO", "NCM", "CST", "CLASS_TRIB", "IBS_UF", "IBS_MUN", "CBS" };
            var faltando = obrigatorias.Where(o => !map.ContainsKey(o)).ToList();
            if (faltando.Any())
                throw new InvalidOperationException($"Colunas obrigatórias não encontradas na planilha: {string.Join(", ", faltando)}");

            return map;
        }

        private void AtualizarBanco(List<TributacaoImportacao> lista)
        {
            var csb = new FbConnectionStringBuilder
            {
                UserID = "SYSDBA",
                Password = "masterkey",
                Database = _caminhoBanco,   // Caminho completo para o arquivo .fdb
                Port = 3050,
                Dialect = 3,
                Charset = "NONE",           // Mude para "NONE" (mais seguro e comum para bancos antigos)
                                            // Ou teste com "WIN1252" se precisar de acentuação correta
                                            // Charset = "WIN1252",
                DataSource = "localhost",   // Adicione isso para conexão TCP local
                ServerType = 0              // 0 = servidor normal (padrão)
            };

            string connectionString = csb.ToString();

            using var con = new FbConnection(connectionString);
            con.Open();
            using var tran = con.BeginTransaction();

            try
            {
                const string sqlComum = """
                    UPDATE PRODUTOS
                    SET CST_IBS_CBS = @CST,
                        CLASS_TRIB = @CLASS,
                        ALIQ_IBS_UF = @ALIQ_IBS_UF,
                        ALIQ_IBS_MUN = @ALIQ_IBS_MUN,
                        ALIQ_CBS = @ALIQ_CBS
                    WHERE 
                """;
                int qtde = 0;
                int count = lista.Count;
                int larguraBarra = 30;
                int rowsAfetts =0;
                foreach (var t in lista)
                {

                    string sql;
                    object parametroChave;
                    string nomeParametro;

                    if (!string.IsNullOrWhiteSpace(t.Produto))
                    {
                        sql = sqlComum + "PRODUTO = @CHAVE";
                        parametroChave = t.Produto.ToString().PadLeft(6, '0');
                        nomeParametro = "@CHAVE";
                    }
                    else if (!string.IsNullOrWhiteSpace(t.Ncm))
                    {
                        sql = sqlComum + "NCM = @CHAVE"; // Ajuste conforme sua tabela real
                        parametroChave = t.Ncm;
                        nomeParametro = "@CHAVE";
                    }
                    else
                    {
                        continue; // Ignora registro sem chave de atualização
                    }

                    using var cmd = new FbCommand(sql, con, tran);
                    cmd.Parameters.Add(nomeParametro, FbDbType.VarChar).Value = parametroChave;
                    cmd.Parameters.Add("@CST", FbDbType.VarChar).Value = (object)t.Cst.ToString().PadLeft(3, '0') ?? DBNull.Value;
                    cmd.Parameters.Add("@CLASS", FbDbType.VarChar).Value = (object)t.ClassTrib.ToString().PadLeft(6, '0') ?? DBNull.Value;
                    cmd.Parameters.Add("@ALIQ_IBS_UF", FbDbType.Double).Value = (double)t.AliqIbsUf;
                    cmd.Parameters.Add("@ALIQ_IBS_MUN", FbDbType.Double).Value = (double)t.AliqIbsMun;
                    cmd.Parameters.Add("@ALIQ_CBS", FbDbType.Double).Value = (double)t.AliqCbs;

                    rowsAfetts += cmd.ExecuteNonQuery();

                    qtde++;
                    double progresso = (double)qtde / count;
                    int barrasPreenchidas = (int)(progresso * larguraBarra);

                    string barra = new string('█', barrasPreenchidas) + new string('░', larguraBarra - barrasPreenchidas);
                    int porcentagem = (int)(progresso * 100);

                    Console.Write($"\rAtualizando: [{barra}] {porcentagem}% ({qtde}/{count}) - Registros afetados: [{rowsAfetts}]");
                   

                }

                tran.Commit();
            }
            catch
            {
                tran.Rollback();
                throw;
            }
        }

        public void Dispose()
        {
            // Caso tenha recursos não gerenciados no futuro
        }
    }


}