using Dwc.Maker.Borders.Entities.Input;

namespace Dwc.Maker.Services.Interfaces;

public interface ILeitorExcel
{
    TabelaDadosBrutos Ler(string caminho);
}