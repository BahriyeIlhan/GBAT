namespace GBAT.Data
{
    internal class Window
    {
        public string Name { get; set; }
        public string ElementNumber { get; set; }
        public decimal GrossArea { get; set; }
        public decimal NetArea { get; set; }
        public decimal Point { get; set; }
        public bool Rsm { get; set; }
        public bool FacadeReuse { get; set; }
        public bool StructureReuse { get; set; }
        public decimal RsmPoint { get; set; }
        public string RsmText { get; set; }
    }
}