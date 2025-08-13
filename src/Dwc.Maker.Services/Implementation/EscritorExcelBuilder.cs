// EscritorExcelBuilder.cs
using ClosedXML.Excel;
using Dwc.Maker.Borders.Entities.Output;

namespace Dwc.Maker.Services.Implementation;

/// <summary>
/// Implementa o padrão Builder para criar e manipular um arquivo Excel
/// usando a biblioteca ClosedXML.
/// </summary>
public class EscritorExcelBuilder : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly string _outputFile;
    private bool _disposed = false;

    /// <summary>
    /// O construtor inicializa uma nova instância do workbook.
    /// </summary>
    public EscritorExcelBuilder(string outputFile)
    {
        _workbook = new XLWorkbook();
        _outputFile = outputFile;
    }

    /// <summary>
    /// Adiciona a aba "Sampling Events" ao workbook.
    /// </summary>
    /// <param name="eventos">A lista de objetos MapeamentoSamplingEvent.</param>
    /// <returns>A própria instância do builder para permitir o encadeamento de métodos.</returns>
    public EscritorExcelBuilder AddAbaSamplingEvents(List<AbaSamplingEvent> eventos)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(EscritorExcelBuilder));
        
        var ws = _workbook.Worksheets.Add("Sampling Events");

        // Define os cabeçalhos
        ws.Cell(1, 1).Value = "parentEventID";
        ws.Cell(1, 2).Value = "eventID";
        ws.Cell(1, 3).Value = "samplingProtocol";
        ws.Cell(1, 4).Value = "samplingEffort";
        ws.Cell(1, 5).Value = "sampleSizeValue";
        ws.Cell(1, 6).Value = "sampleSizeUnit";
        ws.Cell(1, 7).Value = "eventDate";
        ws.Cell(1, 8).Value = "eventRemarks";
        ws.Cell(1, 9).Value = "continent";
        ws.Cell(1, 10).Value = "county";
        ws.Cell(1, 11).Value = "countryCode";
        ws.Cell(1, 12).Value = "stateProvince";
        ws.Cell(1, 13).Value = "municipality";
        ws.Cell(1, 14).Value = "waterBody";
        ws.Cell(1, 15).Value = "habitat";
        ws.Cell(1, 16).Value = "locationId";
        ws.Cell(1, 17).Value = "minimumElevationInMeters";
        ws.Cell(1, 18).Value = "maximumElevationInMeters";
        ws.Cell(1, 19).Value = "decimalLatitude";
        ws.Cell(1, 20).Value = "decimalLongitude";
        ws.Cell(1, 21).Value = "geodeticDatum";
        ws.Cell(1, 22).Value = "coordinateUncertaintyInMeters";
        ws.Cell(1, 23).Value = "georeferenceRemarks";

        // Define o estilo dos cabeçalhos
        ws.Row(1).Style.Font.Bold = true;
        ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;
        ws.Columns().AdjustToContents();

        // Preenche os dados
        for (int i = 0; i < eventos.Count; i++)
        {
            ws.Cell(i + 2, 1).Value = eventos[i].ParentEventID;
            ws.Cell(i + 2, 2).Value = eventos[i].EventID;
            ws.Cell(i + 2, 3).Value = eventos[i].SamplingProtocol;
            ws.Cell(i + 2, 4).Value = eventos[i].SamplingEffort;
            ws.Cell(i + 2, 5).Value = eventos[i].SampleSizeValue;
            ws.Cell(i + 2, 6).Value = eventos[i].SampleSizeUnit;
            ws.Cell(i + 2, 7).Value = eventos[i].EventDate;
            ws.Cell(i + 2, 8).Value = eventos[i].EventRemarks;
            ws.Cell(i + 2, 9).Value = eventos[i].Continent;
            ws.Cell(i + 2, 10).Value = eventos[i].County;
            ws.Cell(i + 2, 11).Value = eventos[i].CountryCode;
            ws.Cell(i + 2, 12).Value = eventos[i].StateProvince;
            ws.Cell(i + 2, 13).Value = eventos[i].Municipality;
            ws.Cell(i + 2, 14).Value = eventos[i].WaterBody;
            ws.Cell(i + 2, 15).Value = eventos[i].Habitat;
            ws.Cell(i + 2, 16).Value = eventos[i].LocationId;
            ws.Cell(i + 2, 17).Value = eventos[i].MinimumElevationInMeters;
            ws.Cell(i + 2, 18).Value = eventos[i].MaximumElevationInMeters;
            ws.Cell(i + 2, 19).Value = eventos[i].DecimalLatitude;
            ws.Cell(i + 2, 20).Value = eventos[i].DecimalLongitude;
            ws.Cell(i + 2, 21).Value = eventos[i].GeodeticDatum;
            ws.Cell(i + 2, 22).Value = eventos[i].CoordinateUncertaintyInMeters;
            ws.Cell(i + 2, 23).Value = eventos[i].GeoreferenceRemarks;
        }

        return this;
    }

    /// <summary>
    /// Adiciona a aba "Associated Occurrences" ao workbook.
    /// </summary>
    /// <param name="ocorrencias">A lista de objetos MapeamentoSamplingEvent.</param>
    /// <returns>A própria instância do builder para permitir o encadeamento de métodos.</returns>
    public EscritorExcelBuilder AddAbaAssociatedOccurrences(List<AbaAssociatedOccurrences> ocorrencias)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(EscritorExcelBuilder));

        var ws = _workbook.Worksheets.Add("Associated Occurrences");

        // Define os cabeçalhos
        ws.Cell(1, 1).Value = "datasetName";
        ws.Cell(1, 2).Value = "eventID";
        ws.Cell(1, 3).Value = "institutionCode";
        ws.Cell(1, 4).Value = "collectionCode";
        ws.Cell(1, 5).Value = "occurrenceID";
        ws.Cell(1, 6).Value = "basisOfRecord";
        ws.Cell(1, 7).Value = "vernacularName";
        ws.Cell(1, 8).Value = "scientificName";
        ws.Cell(1, 9).Value = "scientificNameAuthorship";
        ws.Cell(1, 10).Value = "kingdom";
        ws.Cell(1, 11).Value = "phylum";
        ws.Cell(1, 12).Value = "class";
        ws.Cell(1, 13).Value = "order";
        ws.Cell(1, 14).Value = "family";
        ws.Cell(1, 15).Value = "genus";
        ws.Cell(1, 16).Value = "taxonRank";
        ws.Cell(1, 17).Value = "identificationQualifier";
        ws.Cell(1, 18).Value = "recordedBy";
        ws.Cell(1, 19).Value = "individualCount";
        ws.Cell(1, 20).Value = "organismQuantityType";
        ws.Cell(1, 21).Value = "establishmentMeans";
        ws.Cell(1, 22).Value = "sex";
        ws.Cell(1, 23).Value = "lifeStage";
        ws.Cell(1, 24).Value = "reproductiveCondition";
        ws.Cell(1, 25).Value = "preparations";
        ws.Cell(1, 26).Value = "dynamicProperties";
        ws.Cell(1, 27).Value = "occurrenceRemarks";
        ws.Cell(1, 28).Value = "associatedMedia";
        ws.Cell(1, 29).Value = "taxonDistribution";
        ws.Cell(1, 30).Value = "endemic";
        ws.Cell(1, 31).Value = "rareSpecies";
        ws.Cell(1, 32).Value = "biologicalInterest";
        ws.Cell(1, 33).Value = "scientificInterest";
        ws.Cell(1, 34).Value = "globalStatusConservation";
        ws.Cell(1, 35).Value = "nationalStatusConservation";
        ws.Cell(1, 36).Value = "localStatusConservation";

        // Define o estilo dos cabeçalhos
        ws.Row(1).Style.Font.Bold = true;
        ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;

        // Preenche os dados
        for (int i = 0; i < ocorrencias.Count; i++)
        {
            ws.Cell(i + 2, 1).Value = ocorrencias[i].DatasetName;
            ws.Cell(i + 2, 2).Value = ocorrencias[i].EventID;
            ws.Cell(i + 2, 3).Value = ocorrencias[i].InstitutionCode;
            ws.Cell(i + 2, 4).Value = ocorrencias[i].CollectionCode;
            ws.Cell(i + 2, 5).Value = ocorrencias[i].OccurrenceID;
            ws.Cell(i + 2, 6).Value = ocorrencias[i].BasisOfRecord;
            ws.Cell(i + 2, 7).Value = ocorrencias[i].VernacularName;
            ws.Cell(i + 2, 8).Value = ocorrencias[i].ScientificName;
            ws.Cell(i + 2, 9).Value = ocorrencias[i].ScientificNameAuthorship;
            ws.Cell(i + 2, 10).Value = ocorrencias[i].Kingdom;
            ws.Cell(i + 2, 11).Value = ocorrencias[i].Phylum;
            ws.Cell(i + 2, 12).Value = ocorrencias[i].Class;
            ws.Cell(i + 2, 13).Value = ocorrencias[i].Order;
            ws.Cell(i + 2, 14).Value = ocorrencias[i].Family;
            ws.Cell(i + 2, 15).Value = ocorrencias[i].Genus;
            ws.Cell(i + 2, 16).Value = ocorrencias[i].TaxonRank;
            ws.Cell(i + 2, 17).Value = ocorrencias[i].IdentificationQualifier;
            ws.Cell(i + 2, 18).Value = ocorrencias[i].RecordedBy;
            ws.Cell(i + 2, 19).Value = ocorrencias[i].IndividualCount;
            ws.Cell(i + 2, 20).Value = ocorrencias[i].OrganismQuantityType;
            ws.Cell(i + 2, 21).Value = ocorrencias[i].EstablishmentMeans;
            ws.Cell(i + 2, 22).Value = ocorrencias[i].Sex;
            ws.Cell(i + 2, 23).Value = ocorrencias[i].LifeStage;
            ws.Cell(i + 2, 24).Value = ocorrencias[i].ReproductiveCondition;
            ws.Cell(i + 2, 25).Value = ocorrencias[i].Preparations;
            ws.Cell(i + 2, 26).Value = ocorrencias[i].DynamicProperties;
            ws.Cell(i + 2, 27).Value = ocorrencias[i].OccurrenceRemarks;
            ws.Cell(i + 2, 28).Value = ocorrencias[i].AssociatedMedia;
            ws.Cell(i + 2, 29).Value = ocorrencias[i].TaxonDistribution;
            ws.Cell(i + 2, 30).Value = ocorrencias[i].Endemic;
            ws.Cell(i + 2, 31).Value = ocorrencias[i].RareSpecies;
            ws.Cell(i + 2, 32).Value = ocorrencias[i].BiologicalInterest;
            ws.Cell(i + 2, 33).Value = ocorrencias[i].ScientificInterest;
            ws.Cell(i + 2, 34).Value = ocorrencias[i].GlobalStatusConservation;
            ws.Cell(i + 2, 35).Value = ocorrencias[i].NationalStatusConservation;
            ws.Cell(i + 2, 36).Value = ocorrencias[i].LocalStatusConservation;
        }

        return this;
    }

    /// <summary>
    /// Salva o workbook em um arquivo no caminho especificado.
    /// </summary>
    /// <param name="caminho">O caminho completo do arquivo para salvar.</param>
    public void Build()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(EscritorExcelBuilder));
        _workbook.SaveAs(_outputFile);
    }

    /// <summary>
    /// Implementa o método Dispose para liberar os recursos do workbook.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing) _workbook.Dispose();
        
        _disposed = true;
    }
}
