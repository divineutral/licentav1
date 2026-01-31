namespace LicentaV1.DataObjects
{
    public class MedieSpecializare
    {
        public string Facultate { get; set; } = string.Empty;
        public string Specializare { get; set; } = string.Empty;
        public double NotaMedie { get; set; }
        public int NumarStudenti { get; set; }
    }
}