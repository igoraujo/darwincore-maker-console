using Dwc.Maker.UseCases.Mapeadores;
using Dwc.Maker.Services.Interfaces;
using Dwc.Maker.Services.Implementation;

namespace Dwc.Maker.UseCases.Processadores;

public class PlanilhaProcessor(ILeitorExcel leitor, MapeadorSamplingEvent mapeador)
{
    private readonly ILeitorExcel _leitor = leitor;
    private readonly MapeadorSamplingEvent _mapeador = mapeador;

    public void Executar(string inputFile, string outputFile)
    {
        var dadosBrutos = _leitor.Ler(inputFile);

        var registros = MapeadorSamplingEvent.MapearSamplingEvents(dadosBrutos.AbaRegistro);
        var novaAbaTestes = _mapeador.MapearAssociatedOccurrences(dadosBrutos.AbaRegistro);

        using var darwinCoreBuilder = new EscritorExcelBuilder(outputFile);
        darwinCoreBuilder.AddAbaSamplingEvents(registros)
               .AddAbaAssociatedOccurrences(novaAbaTestes)
               .Build();
    }
}
