using Dwc.Maker.Borders.Entities.Input;
using Dwc.Maker.Services.Interfaces;
using ExcelDataReader;

namespace Dwc.Maker.Services.Implementation;

public class LeitorExcel : ILeitorExcel
{
    public TabelaDadosBrutos Ler(string caminho)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(caminho, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var result = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true
            }
        });

        var registros = result.Tables["CG_BD_MPQ-C24- Registros"];
        // var novaAbaTeste = result.Tables["Nova-Aba"];// ler de uma nova aba

        if (registros is null)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("⚠️  A aba 'CG_BD_MPQ-C24- Registros' não foi encontrada no arquivo Excel.");
            Console.ResetColor();
        }

        // if (novaAbaTeste is null)
        // {
        //     Console.ForegroundColor = ConsoleColor.DarkYellow;
        //     Console.WriteLine("⚠️  A aba 'Nova-Aba' não foi encontrada no arquivo Excel.");
        //     Console.ResetColor();
        // }

        return new TabelaDadosBrutos(registros!);
    }

}
