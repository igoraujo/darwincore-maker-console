namespace Dwc.Maker.ConsoleApp;

public class MenuHandler(ConfigurationManager configManager)
{
    public static void ExibirMenuPrincipal()
    {
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("===========================================");
        Console.WriteLine("♻️ Bem-vindo ao Sistema de Processamento ♻️");
        Console.WriteLine("===========================================");
        Console.ResetColor();
        Console.WriteLine("Escolha uma opção:");
        Console.WriteLine("▶️ Enter. Iniciar");
        Console.WriteLine("⚙️ F2. Configurações");
        Console.WriteLine("❌ Q. Sair do sistema");
        Console.WriteLine("========================================");
    }

    public void ExibirMenuConfiguracoes()
    {
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("====================");
        Console.WriteLine("⚙️ Configurações ⚙️");
        Console.WriteLine("====================");
        Console.ResetColor();
        Console.WriteLine("Escolha uma opção:");
        Console.WriteLine("📂 1. Configurações da pasta de origem");
        Console.WriteLine("👀 2. Visualizar configurações atuais");
        Console.WriteLine("❌ Q. Sair do sistema");
        Console.WriteLine("========================================");

        var opcao = Console.ReadKey().Key;

        switch (opcao)
        {
            case ConsoleKey.D1:
                Console.WriteLine("\n📁 Informe o novo caminho da pasta:");
                var novoBasePath = Console.ReadLine()?.Trim() ?? string.Empty;

                if (!string.IsNullOrEmpty(novoBasePath) && Directory.Exists(novoBasePath))
                {
                    configManager.AtualizarBasePath(novoBasePath);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Caminho base atualizado com sucesso.");
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Caminho inválido. Tente novamente.");
                    Console.ResetColor();
                }
                break;
            case ConsoleKey.D2:
                Console.WriteLine("\nVisualizando configurações atuais...");
                Console.Write($"📂 Caminho da pasta de origem: ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(configManager.BasePath);
                Console.ResetColor();

                break;
            case ConsoleKey.Q:
            case ConsoleKey.Escape:
                return;
            default:
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Opção inválida. Tente novamente.");
                Console.ResetColor();
                break;
        }

        Console.ReadKey();
    }
}
