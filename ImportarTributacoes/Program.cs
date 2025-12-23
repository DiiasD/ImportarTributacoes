using ImportarTributacoes;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;  // Certifique-se de ter esse using

// Configuração da licença não comercial (ESCOLHA UMA DAS DUAS LINHAS ABAIXO)
ExcelPackage.License.SetNonCommercialPersonal("Seu Nome Completo");  // Para uso pessoal
// ExcelPackage.License.SetNonCommercialOrganization("Nome da Sua Organização");  // Para organização não comercial

try
{
    IConfiguration config = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        .Build();

    string caminhoBanco = config["Banco:Caminho"] ?? throw new InvalidOperationException("Configuração 'Banco:Caminho' não encontrada.");
    string caminhoPlanilha = config["Planilha:Caminho"] ?? throw new InvalidOperationException("Configuração 'Planilha:Caminho' não encontrada.");

    if (!int.TryParse(config["Planilha:LinhaInicial"], out int linhaInicial) || linhaInicial < 1)
        throw new InvalidOperationException("Configuração 'Planilha:LinhaInicial' inválida. Deve ser um número inteiro maior ou igual a 1.");

    if (!File.Exists(caminhoBanco))
        throw new FileNotFoundException("Arquivo do banco de dados não encontrado.", caminhoBanco);

    if (!File.Exists(caminhoPlanilha))
        throw new FileNotFoundException("Arquivo da planilha não encontrado.", caminhoPlanilha);

    using var service = new ImportacaoService(caminhoBanco, caminhoPlanilha, linhaInicial);
    service.Executar();

    Console.WriteLine("\nImportação concluída com sucesso!");
    Console.ReadKey();
}
catch (Exception ex)
{
    string logErro = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}{ex}";
    File.WriteAllText("erro.log", logErro);

    Console.WriteLine("Erro durante a importação.");
    Console.WriteLine("Detalhes salvos em 'erro.log'.");
    Console.WriteLine($"Mensagem: {ex.Message}");
}