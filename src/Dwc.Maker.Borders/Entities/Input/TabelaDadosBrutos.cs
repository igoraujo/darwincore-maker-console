using System.Data;

namespace Dwc.Maker.Borders.Entities.Input;

public class TabelaDadosBrutos
{
    public DataTable AbaRegistro { get; set; }
    // public DataTable NovaAbaTeste { get; set; }

    public TabelaDadosBrutos(DataTable registro)
    {
        AbaRegistro = registro;
    }
}
