using Dwc.Maker.UseCases.Mapeadores;
using Dwc.Maker.UseCases.Processadores;
using Microsoft.Extensions.Configuration;
using Dwc.Maker.ConsoleApp;
using Dwc.Maker.Services.Implementation;

Console.OutputEncoding = System.Text.Encoding.UTF8;

var configManager = new Dwc.Maker.ConsoleApp.ConfigurationManager();
var menuHandler = new MenuHandler(configManager);

// Carregar configurações usando caminho absoluto
var configuration = new ConfigurationBuilder()
    .AddJsonFile(Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json"), optional: false, reloadOnChange: true)
    .Build();

string basePath = configuration["BasePath"] ?? "C:";
string inputFile = basePath;
string outputFile = string.Empty;
string fileName = string.Empty;

while (true)
{
    MenuHandler.ExibirMenuPrincipal();
    var opcao = Console.ReadKey().Key;

    switch (opcao)
    {
        case ConsoleKey.Enter:
            Console.Write("\n✏️ Informe o nome do arquivo de dados brutos: ");
            fileName = Console.ReadLine()?.Trim() ?? string.Empty;
            fileName = fileName.EndsWith(".xlsx") ? fileName : $"{fileName}.xlsx";

            inputFile = Path.Combine(configManager.BasePath, fileName);
            outputFile = $"{inputFile.Replace(".xlsx", string.Empty)}_DarwinCore_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            if (!ValidarArquivo(inputFile)) continue;

            ExibirMensagemSucesso("Arquivo carregado com sucesso.");
            break;

        case ConsoleKey.Q:
        case ConsoleKey.Escape:
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n👋 Obrigado por utilizar o sistema. Até logo!");
            Console.ResetColor();
            return;

        case ConsoleKey.F2:
            menuHandler.ExibirMenuConfiguracoes();
            break;

        default:
            ExibirMensagemErro("Opção inválida. Tente novamente.");
            break;
    }

    if (!string.IsNullOrWhiteSpace(fileName))
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.Write("\n⚠️ Tem certeza que deseja processar o arquivo: ");
        Console.ResetColor();
        Console.Write($"{fileName}");
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.Write($"? (S/N)\n");
        Console.ResetColor();

        var confirmacao = Console.ReadKey().Key;

        if (confirmacao == ConsoleKey.S || confirmacao == ConsoleKey.Y || confirmacao == ConsoleKey.Enter)
        {
            ExibirMensagemSucesso($"Arquivo '{fileName}' carregado com sucesso.");
            ExecutarProcessamento(inputFile, outputFile);
        }
        else
        {
            ExibirMensagemErro("Arquivo não confirmado");
        }
    }
}

// Exibe mensagens de erro
void ExibirMensagemErro(string mensagem, string detalhe = "")
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"\n🚫 {mensagem}");
    if (!string.IsNullOrWhiteSpace(detalhe))
    {
        Console.WriteLine(detalhe);
    }

    LimparPaths();
    Console.ResetColor();
    Console.WriteLine("\n🔡 Pressione qualquer tecla para tentar novamente...");
    Console.ReadKey();
}

void LimparPaths()
{
    inputFile = basePath;
    outputFile = string.Empty;
    fileName = string.Empty;
}

// Exibe mensagens de sucesso
void ExibirMensagemSucesso(string mensagem)
{
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"\n✅ {mensagem}");
    Console.ResetColor();
}

static void ExibirBarraDeProgresso(int delayMs)
{
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine("🔄 Processando...\n");
    Console.ResetColor();

    int larguraBarra = 30; // número de blocos visuais

    for (int i = 0; i <= larguraBarra; i++)
    {
        double percentual = i / (double)larguraBarra * 100;

        // Monta a barra usando blocos cheios
        string barra = new string('█', i).PadRight(larguraBarra, ' ');

        Console.ForegroundColor = ConsoleColor.Green;
        // Reescreve a linha com a barra e percentual
        Console.Write($"\r{barra}");
        Console.ResetColor();
        Console.Write($" {percentual:0}%");

        Thread.Sleep(delayMs);
    }
    Console.ResetColor();
    Console.WriteLine(" \n\n✅ Concluído!");
}


// Executa o processamento
void ExecutarProcessamento(string inputFile, string outputFile)
{
    try
    {
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine("\n🔄 Iniciando o processamento da planilha...");
        Console.ResetColor();

        var leitor = new LeitorExcel();
        var mapeador = new MapeadorSamplingEvent();
        var processor = new PlanilhaProcessor(leitor, mapeador);
        processor.Executar(inputFile, outputFile);
        ExibirBarraDeProgresso(10);
        ExibirMensagemSucesso($"Processamento concluído.");
        Console.Write($"\nArquivo exportado em:");
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.Write($"{outputFile}");
        Console.ResetColor();


    }
    catch (Exception ex)
    {
        ExibirMensagemErro("Erro ao processar planilha:", ex.Message);
    }

    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine("\n🔡 Pressione qualquer tecla para voltar ao menu ou 'Esc' para sair.");
    Console.ResetColor();
    var key = Console.ReadKey();

    if (key.Key == ConsoleKey.Escape || key.Key == ConsoleKey.Q) return;

}

// Corrigir acesso à função ValidarArquivo
bool ValidarArquivo(string caminho)
{
    if (!File.Exists(caminho))
    {
        ExibirMensagemErro($"Arquivo '{caminho}' não encontrado.");
        return false;
    }
    return true;
}