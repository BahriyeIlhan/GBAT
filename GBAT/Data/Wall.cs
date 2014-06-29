namespace GBAT.Data
{
    internal class Wall
    {
        public WallPositionType WallPositionType { get; set; }
        public string Name { get; set; }
        public string ElementNumber { get; set; }
        public decimal GrossFootprintArea { get; set; }
        public decimal NetFootprintArea { get; set; }
        public decimal GrossSideArea { get; set; }
        public decimal NetSideArea { get; set; }
        public decimal Point { get; set; }
        public bool Rsm { get; set; }
        public decimal InsulationLayerThickness { get; set; }
        public decimal InsulationLayerMaterialPoint { get; set; }
        public string InsulationMaterialName { get; set; }
        public string InsulationMaterialElementNumber { get; set; }
        public bool FacadeReuse { get; set; }
        public bool StructureReuse { get; set; }
        public decimal RsmPoint { get; set; }
        public string RsmText { get; set; }
        public decimal NetVolume { get; set; }
        public bool LoadBearing { get; set; }
    }
}