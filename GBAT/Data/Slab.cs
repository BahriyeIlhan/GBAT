namespace GBAT.Data
{
    internal class Slab
    {
        public string Name { get; set; }
        public string ElementNumber { get; set; }
        public SlabType SlabType { get; set; }
        public decimal GrossArea { get; set; }
        public decimal NetArea { get; set; }
        public decimal Point { get; set; }
        public decimal InsulationLayerThickness { get; set; }
        public decimal InsulationLayerMaterialPoint { get; set; }
        public string InsulationMaterialName { get; set; }
        public string InsulationMaterialElementNumber { get; set; }
        public bool Rsm { get; set; }
        public bool FacadeReuse { get; set; }
        public bool StructureReuse { get; set; }
        public decimal RsmPoint { get; set; }
        public string RsmText { get; set; }
        public decimal NetVolume { get; set; }
        public bool LoadBearing { get; set; }
    }
}