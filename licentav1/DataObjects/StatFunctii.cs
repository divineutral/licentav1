namespace LicentaV1.DataObjects
{
    public class StatFunctii
    {
        public string TipPost { get; set; } = string.Empty;
        public int TitularOcupate { get; set; }
        public int PlataCuOra { get; set; }
        public int Vacante { get; set; }
        public int Total => TitularOcupate + PlataCuOra + Vacante;
    }
}