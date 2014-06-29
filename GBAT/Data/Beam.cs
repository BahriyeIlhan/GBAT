namespace GBAT.Data
{
    internal class Beam
    {
        public string Name { get; set; }
        public string ElementNumber { get; set; }
        public decimal CrossSectionArea { get; set; }
        public decimal OuterSurfaceArea { get; set; }
        public decimal TotalSurfaceArea { get; set; }
        public decimal GrossSurfaceArea { get; set; }
        public decimal Point { get; set; }
        public bool Rsm { get; set; }
        public bool FacadeReuse { get; set; }
        public bool StructureReuse { get; set; }
        public decimal RsmPoint { get; set; }
        public string RsmText { get; set; }
        public decimal NetSurfaceAreaExtrudedSides { get; set; }
        public decimal NetVolume { get; set; }
    }
}