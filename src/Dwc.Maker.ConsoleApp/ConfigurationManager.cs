using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace Dwc.Maker.ConsoleApp;

public class ConfigurationManager
{
    private readonly IConfiguration _configuration;
    private readonly string _configFilePath;

    public string BasePath { get; private set; }

    public ConfigurationManager()
    {
        _configFilePath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
        _configuration = new ConfigurationBuilder()
            .AddJsonFile(_configFilePath, optional: false, reloadOnChange: true)
            .Build();

        BasePath = _configuration["BasePath"] ?? string.Empty;
    }

    public void AtualizarBasePath(string novoBasePath)
    {
        BasePath = novoBasePath;
        SalvarConfiguracao("BasePath", novoBasePath);
    }

    private void SalvarConfiguracao(string chave, string valor)
    {
        var json = File.ReadAllText(_configFilePath);
        var jsonObj = JsonConvert.DeserializeObject<Dictionary<string, string>>(json) ?? [];
        jsonObj[chave] = valor;
        File.WriteAllText(_configFilePath, JsonConvert.SerializeObject(jsonObj, Formatting.Indented));
    }
}
