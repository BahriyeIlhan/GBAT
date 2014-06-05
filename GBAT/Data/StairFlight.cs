namespace GBAT.Data
{
    internal class StairFlight
    {
        public int NumberOfRiser { get; set; }
        public int NumberOfTreads { get; set; }
        public decimal RiserHeight { get; set; }
        public decimal TreadLength { get; set; }
        public decimal TreadDistance { get; set; }
        public decimal Area
        {
            get
            {
                return (NumberOfRiser * RiserHeight * TreadDistance) + (NumberOfTreads * TreadLength * TreadDistance);
            }
        }
        public bool Rsm { get; set; }
        public bool StructureReuse { get; set; }
        public bool FacadeReuse { get; set; }
        public decimal RsmPoint { get; set; }
        public string RsmText { get; set; }
        public string Name { get; set; }
        public string ElementNumber { get; set; }
    }
}