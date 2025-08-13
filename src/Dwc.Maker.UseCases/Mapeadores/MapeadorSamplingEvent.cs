using System.Data;
using Dwc.Maker.Borders.Dtos;
using Dwc.Maker.Borders.Entities.Output;
using Dwc.Maker.Borders.Enums;

namespace Dwc.Maker.UseCases.Mapeadores;

public class MapeadorSamplingEvent
{
    public static List<AbaSamplingEvent> MapearSamplingEvents(DataTable registros)
    {
        if (registros is null || registros.Rows.Count <= 2)
        {
            MostrarErro("A aba de registros está vazia ou incompleta!");
            return [];
        }

        var lista = new List<AbaSamplingEvent>();
        AbaInputRegistroDto abaRegistro = new();

        for (int i = 2; i < registros.Rows.Count; i++) // Pula cabeçalhos
        {
            MontarAbaRegistro(registros, ref abaRegistro, i);

            if (string.IsNullOrWhiteSpace(abaRegistro.Campanha) || string.IsNullOrWhiteSpace(abaRegistro.Estacao))
                continue;
            string numeroFormatadoEventId = i.ToString("D6");

            lista.Add(new AbaSamplingEvent
            {
                ParentEventID = $"C{abaRegistro.Campanha}:{abaRegistro.Estacao}",
                EventID = $"C{abaRegistro.Campanha}:{abaRegistro.Estacao}:mamm:{numeroFormatadoEventId}",
                SamplingProtocol = abaRegistro.CapturaMetodo,
                SamplingEffort = "80 armadilhas-noite",
                SampleSizeValue = "24",
                SampleSizeUnit = "hora",
                EventDate = "-",
                EventRemarks = abaRegistro.CapturaMetodo,
                Continent = "South America",
                County = "Brazil",
                CountryCode = "BR",
                StateProvince = "Minas Gerais",
                Municipality = abaRegistro.Municipio,
                WaterBody = "-",
                Habitat = "-",
                LocationId = abaRegistro.Area,
                MinimumElevationInMeters = "-",
                MaximumElevationInMeters = "-",
                DecimalLatitude = abaRegistro.CoordLat,
                DecimalLongitude = abaRegistro.CoordLong,
                GeodeticDatum = "SIRGAS, 2000",
                CoordinateUncertaintyInMeters = string.Empty,
                GeoreferenceRemarks = string.Empty,
            });
        }

        return lista;
    }

    public List<AbaAssociatedOccurrences> MapearAssociatedOccurrences(DataTable registros)
    {
        if (registros is null || registros.Rows.Count <= 2)
        {
            MostrarErro("A aba de registros está vazia ou incompleta!");
            return [];
        }

        var lista = new List<AbaAssociatedOccurrences>();
        AbaInputRegistroDto abaRegistro = new();

        for (int i = 2; i < registros.Rows.Count; i++) // Pula cabeçalhos
        {
            MontarAbaRegistro(registros, ref abaRegistro, i);

            if (string.IsNullOrWhiteSpace(abaRegistro.Campanha) || string.IsNullOrWhiteSpace(abaRegistro.Estacao))
                continue;
            string numeroFormatadoEventId = i.ToString("D6");

            lista.Add(new AbaAssociatedOccurrences
            {
                DatasetName = "Samarco-DataSet-bio",
                EventID = $"C{abaRegistro.Campanha}:{abaRegistro.Estacao}:mamm:{numeroFormatadoEventId}",
                InstitutionCode = "SAM",
                CollectionCode = "mamm",
                OccurrenceID = "https://www.gbif.org/species/",
                BasisOfRecord = "-",
                VernacularName = abaRegistro.NomePopular,
                ScientificName = abaRegistro.Especie,
                ScientificNameAuthorship = $"-",
                Kingdom = "Animalia",
                Phylum = "-",
                Class = "-",
                Order = abaRegistro.Ordem,
                Family = abaRegistro.Familia,
                Genus = abaRegistro.Genero,
                TaxonRank = "species",
                IdentificationQualifier = "-",
                RecordedBy = abaRegistro.Coletor,
                IndividualCount = "1",
                OrganismQuantityType = "individuals",
                EstablishmentMeans = "-",
                Sex = abaRegistro.Sexo,
                LifeStage = abaRegistro.ClasseEtaria,
                ReproductiveCondition = "-",
                Preparations = "-",
                DynamicProperties = abaRegistro.Iucn2024,
                OccurrenceRemarks = abaRegistro.Observacoes,
                AssociatedMedia = "-",
                TaxonDistribution = "American continent",
                Endemic = abaRegistro.Endemica,
                RareSpecies = "-",
                BiologicalInterest = "-",
                ScientificInterest = "-",
                GlobalStatusConservation = abaRegistro.Iucn2024,
                NationalStatusConservation = abaRegistro.Mma2022,
                LocalStatusConservation = abaRegistro.Copam1472010,
            });
        }

        return lista;
    }

    private static void MontarAbaRegistro(DataTable registros, ref AbaInputRegistroDto abaRegistro, int i)
    {
        var linha = registros.Rows[i];

        abaRegistro.Registro = ObterValor(linha, DadosBrutosColuna.Registro);
        abaRegistro.Area = ObterValor(linha, DadosBrutosColuna.Area);
        abaRegistro.CoordLat = registros.Rows[i][(int)DadosBrutosColuna.CoordLat]?.ToString();
        abaRegistro.CoordLong = ObterValor(linha, DadosBrutosColuna.CoordLong);
        abaRegistro.Estado = ObterValor(linha, DadosBrutosColuna.Estado);
        abaRegistro.Municipio = ObterValor(linha, DadosBrutosColuna.Municipio);
        abaRegistro.Ano = ObterValor(linha, DadosBrutosColuna.Ano);
        abaRegistro.Mes = ObterValor(linha, DadosBrutosColuna.Mes);
        abaRegistro.Dia = ObterValor(linha, DadosBrutosColuna.Dia);
        abaRegistro.Campanha = ObterValor(linha, DadosBrutosColuna.Campanha);
        abaRegistro.Estacao = ObterValor(linha, DadosBrutosColuna.Estacao);
        abaRegistro.Coletor = ObterValor(linha, DadosBrutosColuna.Coletor);
        abaRegistro.CapturaMetodo = ObterValor(linha, DadosBrutosColuna.CapturaMetodo);
        abaRegistro.TipoRegistro = ObterValor(linha, DadosBrutosColuna.TipoRegistro);
        abaRegistro.Posto = ObterValor(linha, DadosBrutosColuna.Posto);
        abaRegistro.Estrato = ObterValor(linha, DadosBrutosColuna.Estrato);
        abaRegistro.Anilha = ObterValor(linha, DadosBrutosColuna.Anilha);
        abaRegistro.Captura = ObterValor(linha, DadosBrutosColuna.Captura);
        abaRegistro.Ordem = ObterValor(linha, DadosBrutosColuna.Ordem);
        abaRegistro.Familia = ObterValor(linha, DadosBrutosColuna.Familia);
        abaRegistro.Genero = ObterValor(linha, DadosBrutosColuna.Genero);
        abaRegistro.Especie = ObterValor(linha, DadosBrutosColuna.Especie);
        abaRegistro.NomePopular = ObterValor(linha, DadosBrutosColuna.NomePopular);
        abaRegistro.Destino = ObterValor(linha, DadosBrutosColuna.Destino);
        abaRegistro.CompCorpoMM = ObterValor(linha, DadosBrutosColuna.CompCorpoMM);
        abaRegistro.CompCaudaMM = ObterValor(linha, DadosBrutosColuna.CompCaudaMM);
        abaRegistro.CompTarsoMM = ObterValor(linha, DadosBrutosColuna.CompTarsoMM);
        abaRegistro.CompOrelhaMM = ObterValor(linha, DadosBrutosColuna.CompOrelhaMM);
        abaRegistro.PesoSacoArmAnimalG = ObterValor(linha, DadosBrutosColuna.PesoSacoArmAnimalG);
        abaRegistro.PesoSacoArmadG = ObterValor(linha, DadosBrutosColuna.PesoSacoArmadG);
        abaRegistro.PesoAnimalG = ObterValor(linha, DadosBrutosColuna.PesoAnimalG);
        abaRegistro.Sexo = ObterValor(linha, DadosBrutosColuna.Sexo);
        abaRegistro.ClasseEtaria = ObterValor(linha, DadosBrutosColuna.ClasseEtaria);
        abaRegistro.Observacoes = ObterValor(linha, DadosBrutosColuna.Observacoes);
        abaRegistro.Iucn2024 = ObterValor(linha, DadosBrutosColuna.IUCN2024);
        abaRegistro.Mma2022 = ObterValor(linha, DadosBrutosColuna.MMA2022);
        abaRegistro.Copam1472010 = ObterValor(linha, DadosBrutosColuna.COPAM1472010);
        abaRegistro.Endemica = ObterValor(linha, DadosBrutosColuna.Endemica);
        abaRegistro.DadosInsuficientes = ObterValor(linha, DadosBrutosColuna.DadosInsuficientes);
        abaRegistro.ExoticaInvasora = ObterValor(linha, DadosBrutosColuna.ExoticaInvasora);
    }

    private static string ObterValor(DataRow row, DadosBrutosColuna coluna) =>
        row[(int)coluna]?.ToString()?.Trim() ?? string.Empty;


    private static void MostrarErro(string mensagem)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"🚫 {mensagem}");
        Console.ResetColor();
    }
}
