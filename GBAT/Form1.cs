using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using GBAT.Data;
using sharpPDF;
using sharpPDF.Enumerators;
using sharpPDF.Tables;

namespace GBAT
{
	public partial class Form1 : Form
	{
		private Process process;
		private DataSet dataset;
		private DataTable materialTable;
		private const decimal Conductivity = (decimal)0.04;
		private const string IdColumn = "Column1";
		private string buildingType;
		private string projectName;
		private string projectType;
		private string projectPhase;
		private Dictionary<string, string> documentCertficate;

		private List<Wall> externalWalls;
		private List<Wall> interiorWalls;
		private List<Wall> curtainWalls;
		private List<Wall> boundaryWalls;
		private List<Window> windows;
		private List<Slab> slabs;
		private List<Column> columns;
		private List<Beam> beams;
		private List<StairFlight> stairs;

		public Form1()
		{
			InitializeComponent();
			PopulateMaterailList();
		}

		private void PopulateMaterailList()
		{
			MaterialList.ValueMember = "Value";
			MaterialList.DisplayMember = "Text";
			MaterialList.Items.Add(new { Text = "Mat 1: Material Specification", Value = "Mat1" });
			MaterialList.Items.Add(new { Text = "Mat 2a: Natural Boundary", Value = "Mat2a" });
			MaterialList.Items.Add(new { Text = "Mat 2b: Hard Landscaping and Boundary Protection", Value = "Mat2b" });
			MaterialList.Items.Add(new { Text = "Mat 3: Facade Reuse", Value = "Mat3" });
			MaterialList.Items.Add(new { Text = "Mat 4: Structure Reuse", Value = "Mat4" });
			MaterialList.Items.Add(new { Text = "Mat 5: Responsibly Sourced Materials", Value = "Mat5" });
			MaterialList.Items.Add(new { Text = "Mat 6a: Insulation", Value = "Mat6a" });
			MaterialList.Items.Add(new { Text = "Mat 6b: Responsibly Sourced Insulation", Value = "Mat6b" });
			MaterialList.Items.Add(new { Text = "Mat 7: Designing for Robustness", Value = "Mat7" });
		}

		private void BrowseIfcFileClicked(object sender, EventArgs e)
		{
			var openFileDialog = new OpenFileDialog { DefaultExt = ".ifc", AddExtension = true };
			var dialogResult = openFileDialog.ShowDialog();
			if (dialogResult == DialogResult.OK)
			{
				IFCFileUrl.Text = openFileDialog.FileName;
			}
		}

		private void ProcessIfcFileClicked(object sender, EventArgs e)
		{
			var ifcFileUrl = IFCFileUrl.Text;
			if (string.IsNullOrEmpty(ifcFileUrl))
			{
				MessageBox.Show(@"IFC file must be selected");
				return;
			}
			if (!File.Exists(ifcFileUrl))
			{
				MessageBox.Show(@"IFC file not found");
				return;
			}

			var xlsxFileName = ifcFileUrl.Replace(".ifc", "_ifc") + ".xlsx";
			if (File.Exists(xlsxFileName))
			{
				MainProcess();
			}
			else
			{
				var processInfo = new ProcessStartInfo("IFC Analyzer\\IFC-File-Analyzer-CL.exe", string.Format("{0} {1}", ifcFileUrl, "noopen"))
				{
					CreateNoWindow = true,
					UseShellExecute = false
				};
				CancelProcess.Enabled = true;
				Results.Text = @"Started";
				IFAProgress.Visible = true;
				process = Process.Start(processInfo);
				process.EnableRaisingEvents = true;
				process.Exited += ProcessOnExited;
			}

		}

		private Dictionary<string, string> CalculateCertification(string projectType, string buildingType)
		{
			var selectedMaterials = RetrieveSelectedMaterials();
			var mat1 = RetrievePropertyNominalValueByName("Mat 1_Material Specification");
			var mat2A = RetrievePropertyNominalValueByName("Mat 2a_Natural Boundary");
			var mat2B = RetrievePropertyNominalValueByName("Mat 2b_Hard Landscaping and Boundary Protection");
			var mat3 = RetrievePropertyNominalValueByName("Mat 3_Facade Reuse");
			var mat4 = RetrievePropertyNominalValueByName("Mat 4_Structure Reuse");
			var mat5 = RetrievePropertyNominalValueByName("Mat 5_Responsibly Sourced Materials");
			var mat6A = RetrievePropertyNominalValueByName("Mat 6a_Insulation");
			var mat6B = RetrievePropertyNominalValueByName("Mat 6b_Responsibly Sourced Insulation");
			var mat7 = RetrievePropertyNominalValueByName("Mat 7_Designing for Robustness");

			this.externalWalls = RetrieveWallsByPosition(WallPositionType.External);
			this.interiorWalls = RetrieveWallsByPosition(WallPositionType.Interior);
			this.curtainWalls = RetrieveCurtainWalls();
			this.boundaryWalls = RetrieveWallsByPosition(WallPositionType.Boundary);
			this.windows = RetrieveWindows();
			this.slabs = RetrieveSlabs();
			this.columns = RetrieveColumns();
			this.stairs = RetrieveStairs();
			this.beams = RetrieveBeams();

			var certification = new Dictionary<string, string>();
			documentCertficate = new Dictionary<string, string>();

			#region mat1
			if (mat1 == "1" && selectedMaterials.Contains("Mat1"))
			{
				var mat1WallTotalPoint = ((externalWalls.Sum(ew => ew.NetSideArea * ew.Point))
									 + (curtainWalls.Sum(cw => cw.GrossSideArea * cw.Point)));

				var mat1WindowTotalPoint = windows.Sum(w => w.NetArea * w.Point);

				var mat1RoofTotalPoint = slabs.Where(s => s.SlabType == SlabType.Roof).Sum(s => s.NetArea * s.Point);
				var mat1FloorTotalPoint = slabs.Where(s => s.SlabType == SlabType.Floor).Sum(s => s.NetArea * s.Point);
				//var mat1SlabTotalPoint = slabs.Where(s => s.SlabType == SlabType.Floor || s.SlabType == SlabType.Roof).Sum(s => s.NetArea * s.Point);
				var mat1SlabTotalPoint = mat1RoofTotalPoint + mat1FloorTotalPoint;
				var mat1TotalArea = externalWalls.Sum(ew => ew.NetSideArea)
									+ curtainWalls.Sum(cw => cw.GrossSideArea) +
									windows.Sum(w => w.NetArea) + slabs.Where(s => s.SlabType == SlabType.Floor || s.SlabType == SlabType.Roof).Sum(s => s.NetArea);

				var mat1Point = (mat1WallTotalPoint + mat1WindowTotalPoint + mat1SlabTotalPoint) / (mat1TotalArea);

				var mat1Certification = CalculateMatOneCertification(mat1Point, projectType, buildingType);
				certification.Add("Mat1: Material Specification", mat1Certification);
				documentCertficate.Add("Mat1", mat1Certification);
			}
			else if (selectedMaterials.Contains("Mat1"))
			{
				certification.Add("Mat1: Material Specification", "No Credit");
				documentCertficate.Add("Mat1", "No Credit");
			}

			#endregion

			#region mat2
			if (mat2A == "1" && selectedMaterials.Contains("Mat2a"))
			{
				certification.Add("Mat2a: Natural Boundary", "1 Credit");
				documentCertficate.Add("Mat2a", "1 Credit");
			}
			else if (selectedMaterials.Contains("Mat2a"))
			{
				certification.Add("Mat2a: Natural Boundary", "No Credit");
				documentCertficate.Add("Mat2a", "No Credit");
			}

			if (mat2B == "1" && selectedMaterials.Contains("Mat2b") && (!documentCertficate.ContainsKey("Mat2a") || documentCertficate["Mat2a"] == "No Credit"))
			{
				var mat2WallTotalPoint = boundaryWalls.Sum(bw => bw.NetSideArea * bw.Point);
				var landingSlabs = slabs.Where(s => s.SlabType == SlabType.Landing);
				var mat2SlabTotalPoint = landingSlabs.Sum(ls => ls.NetArea * ls.Point);
				var mat2TotalArea = boundaryWalls.Sum(bw => bw.NetSideArea) + landingSlabs.Sum(ls => ls.NetArea);
				var mat2Point = (mat2WallTotalPoint + mat2SlabTotalPoint) / mat2TotalArea;
				if (mat2Point / (decimal)0.8 >= 2)
				{
					certification.Add("Mat2b: Hard Landscaping", "1 Credit");
					documentCertficate.Add("Mat2b", "1 Credit");
				}
				else
				{
					certification.Add("Mat2b: Hard Landscaping", "0 Credit");
					documentCertficate.Add("Mat2b", "0 Credit");
				}
			}
			#endregion

			#region mat3
			if (mat3 == "1" && selectedMaterials.Contains("Mat3"))
			{
				var totalArea = externalWalls.Sum(ew => ew.NetSideArea)
								+ curtainWalls.Sum(cw => cw.GrossSideArea)
								+ windows.Sum(w => w.NetArea);
				var facadeReuseArea = externalWalls.Where(e => e.FacadeReuse).Sum(ew => ew.NetSideArea)
									  + curtainWalls.Where(e => e.FacadeReuse).Sum(cw => cw.GrossSideArea)
									  + windows.Where(e => e.FacadeReuse).Sum(w => w.NetArea);
				if (facadeReuseArea / totalArea >= (decimal)0.5)
				{
					certification.Add("Mat3: Facade Reuse", "1 Credit");
					documentCertficate.Add("Mat3", "1 Credit");
				}
				else
				{
					certification.Add("Mat3: Facade Reuse", "0 Credit");
					documentCertficate.Add("Mat3", "0 Credit");
				}
			}
			else if (selectedMaterials.Contains("Mat3"))
			{
				certification.Add("Mat3: Facade Reuse", "No Credit");
				documentCertficate.Add("Mat3", "No Credit");
			}
			#endregion

			#region mat4
			if (mat4 == "1" && selectedMaterials.Contains("Mat4"))
			{
				if (!projectType.Contains("Major Refurbishment"))
				{
					certification.Add("Mat4: Structure Reuse", "0 Credit");
				}
				else
				{
					var mat4TotalVolume = externalWalls.Where(ew => ew.LoadBearing).Sum(ew => ew.NetVolume) +
										  interiorWalls.Where(iw => iw.LoadBearing).Sum(iw => iw.NetVolume)
										  +
										  slabs.Where(s => s.SlabType == SlabType.Floor && s.LoadBearing)
											  .Sum(s => s.NetVolume) + columns.Sum(c => c.NetVolume) + beams.Sum(b => b.NetVolume);

					var mat4StructureReusedVolume = externalWalls.Where(ew => ew.LoadBearing && ew.StructureReuse).Sum(ew => ew.NetVolume) +
										  interiorWalls.Where(iw => iw.LoadBearing && iw.StructureReuse).Sum(iw => iw.NetVolume)
										  +
										  slabs.Where(s => s.SlabType == SlabType.Floor && s.LoadBearing && s.StructureReuse)
											  .Sum(s => s.NetVolume)
											  + columns.Where(c => c.StructureReuse).Sum(c => c.NetVolume) + beams.Where(b => b.StructureReuse).Sum(b => b.NetVolume);

					if (mat4StructureReusedVolume / mat4TotalVolume >= (decimal)0.8)
					{
						certification.Add("Mat4: Structure Reuse", "1 Credit");
						documentCertficate.Add("Mat4", "1 Credit");
					}
					else
					{
						certification.Add("Mat4: Structure Reuse", "0 Credit");
						documentCertficate.Add("Mat4", "0 Credit");
					}
				}

			}
			else if (selectedMaterials.Contains("Mat4"))
			{
				certification.Add("Mat4: Structure Reuse", "No Credit");
				documentCertficate.Add("Mat4", "No Credit");
			}
			#endregion

			#region mat5
			if (mat5 == "1" && selectedMaterials.Contains("Mat5"))
			{
				if (!projectType.Contains("Fit-Out"))
				{
					var totalRsmArea = (decimal)0;
					var totalPoint = (decimal)0;

					var mat5ExternalWallPoint = (decimal)0;

					var mat5ExternalWallTotalArea = externalWalls.Sum(ew => ew.NetSideArea);
					var mat5ExternalWallRsmTotalArea = externalWalls.Where(e => e.Rsm).Sum(ew => ew.NetSideArea);
					if (mat5ExternalWallRsmTotalArea / mat5ExternalWallTotalArea >= (decimal)0.8)
					{
						mat5ExternalWallPoint = externalWalls.Where(e => e.Rsm).Sum(ew => ew.NetSideArea * ew.RsmPoint) / mat5ExternalWallRsmTotalArea;
					}

					var mat5InteriorWallTotalArea = interiorWalls.Sum(iw => iw.NetSideArea);
					var mat5InteriorWallRsmTotalArea = interiorWalls.Where(e => e.Rsm).Sum(iw => iw.NetSideArea);
					var mat5InteriorWallPoint = (decimal)0;
					if (mat5InteriorWallRsmTotalArea / mat5InteriorWallTotalArea >= (decimal)0.8)
					{
						mat5InteriorWallPoint = interiorWalls.Where(e => e.Rsm).Sum(iw => iw.NetSideArea * iw.RsmPoint) /
												mat5InteriorWallRsmTotalArea;
					}

					var mat5RoofSlabPoint = (decimal)0;
					var mat5RoofSlabTotalArea = slabs.Where(e => e.SlabType == SlabType.Roof).Sum(s => s.NetArea);
					var mat5RoofSlabRsmTotalArea =
						slabs.Where(e => e.Rsm && e.SlabType == SlabType.Roof).Sum(s => s.NetArea);

					if (mat5RoofSlabRsmTotalArea / mat5RoofSlabTotalArea >= (decimal)0.8)
					{
						mat5RoofSlabPoint =
							slabs.Where(e => e.Rsm && e.SlabType == SlabType.Roof).Sum(s => s.NetArea * s.RsmPoint) /
							mat5RoofSlabRsmTotalArea;
					}

					var mat5UpperFloorPoint = (decimal)0;
					var mat5UpperFloorTotalArea = slabs.Where(e => e.SlabType == SlabType.Floor).Sum(s => s.NetArea);
					var mat5UpperFloorRsmTotalArea =
						slabs.Where(e => e.SlabType == SlabType.Floor && e.Rsm).Sum(s => s.NetArea);
					if (mat5UpperFloorRsmTotalArea / mat5UpperFloorTotalArea >= (decimal)0.8)
					{
						mat5UpperFloorPoint =
							slabs.Where(e => e.SlabType == SlabType.Floor && e.Rsm).Sum(s => s.NetArea * s.RsmPoint) /
							mat5UpperFloorRsmTotalArea;
					}

					var mat5BaseSlabPoint = (decimal)0;
					var mat5BaseSlabTotalArea = slabs.Where(s => s.SlabType == SlabType.BaseSlab).Sum(s => s.NetArea);
					var mat5BaseSlabRsmTotalArea =
						slabs.Where(s => s.SlabType == SlabType.BaseSlab && s.Rsm).Sum(s => s.NetArea);
					if (mat5BaseSlabRsmTotalArea / mat5BaseSlabTotalArea > (decimal)0.8)
					{
						mat5BaseSlabPoint =
							slabs.Where(e => e.SlabType == SlabType.BaseSlab && e.Rsm).Sum(s => s.NetArea * s.RsmPoint) /
							mat5BaseSlabRsmTotalArea;
					}

					var mat5ColumnTotalArea = columns.Sum(c => c.GrossSurfaceArea);
					var mat5ColumnRsmTotalArea = columns.Where(e => e.Rsm).Sum(c => c.GrossSurfaceArea);
					var mat5ColumnPoint = (decimal)0;
					if (mat5ColumnRsmTotalArea / mat5ColumnTotalArea >= (decimal)0.8)
					{
						mat5ColumnPoint = columns.Where(c => c.Rsm).Sum(c => c.GrossSurfaceArea * c.RsmPoint) /
										  mat5ColumnRsmTotalArea;
					}

					var mat5BeamTotalArea = beams.Sum(b => b.GrossSurfaceArea);
					var mat5BeamRsmTotalArea = beams.Where(b => b.Rsm).Sum(b => b.GrossSurfaceArea);
					var mat5BeamPoint = (decimal)0;
					if (mat5BeamRsmTotalArea / mat5BeamTotalArea >= (decimal)0.8)
					{
						mat5BeamPoint = beams.Where(b => b.Rsm).Sum(b => b.NetSurfaceAreaExtrudedSides * b.RsmPoint) /
										mat5BeamRsmTotalArea;
					}

					var mat5StairTotalArea = stairs.Sum(s => s.Area);
					var mat5StairRsmTotalArea = stairs.Where(e => e.Rsm).Sum(s => s.Area);
					var mat5StairPoint = (decimal)0;
					if (mat5StairRsmTotalArea / mat5StairTotalArea >= (decimal)0.8)
					{
						mat5StairPoint = stairs.Where(e => e.Rsm).Sum(s => s.Area * s.RsmPoint) / mat5StairRsmTotalArea;
					}

					totalPoint = mat5ExternalWallPoint + mat5InteriorWallPoint + mat5BaseSlabPoint + mat5BeamPoint +
								 mat5ColumnPoint + mat5RoofSlabPoint + mat5StairPoint + mat5UpperFloorPoint;

					if (totalPoint == 0)
					{
						certification.Add("Mat5: Responisbly Sourced Materials", "0 Credit");
						documentCertficate.Add("Mat5", "0 Credit");
					}


					else if (totalPoint < 5)
					{
						certification.Add("Mat5: Responisbly Sourced Materials", "0 Credit");
						documentCertficate.Add("Mat5", "0 Credit");
					}
					else if (totalPoint >= 5 && totalPoint < 10)
					{
						certification.Add("Mat5: Responisbly Sourced Materials", "1 Credit");
						documentCertficate.Add("Mat5", "1 Credit");
					}
					else if (totalPoint >= 10 && totalPoint < 15)
					{
						certification.Add("Mat5: Responisbly Sourced Materials", "2 Credit");
						documentCertficate.Add("Mat5", "2 Credit");
					}
					else
					{
						certification.Add("Mat5: Responisbly Sourced Materials", "3 Credit");
						documentCertficate.Add("Mat5", "3 Credit");
					}
				}
			}
			else if (selectedMaterials.Contains("Mat5"))
			{
				certification.Add("Mat5: Responisbly Sourced Materials", "No Credit");
				documentCertficate.Add("Mat5", "No Credit");
			}
			#endregion

			#region mat6

			if (mat6A == "1" && selectedMaterials.Contains("Mat6a"))
			{
				var mat6AWallThermalResistance =
					externalWalls.Sum(
						ew =>
							(ew.NetSideArea * (ew.InsulationLayerThickness / 1000) / Conductivity));

				var mat6AWallGreenGuideCorrected = externalWalls.Sum(
						ew => (ew.NetSideArea * (ew.InsulationLayerThickness / 1000) / Conductivity) * ew.InsulationLayerMaterialPoint);

				var mat6ASlabThermalResistance = slabs.Where(s => s.SlabType == SlabType.BaseSlab || s.SlabType == SlabType.Roof).Sum(s => (s.NetArea * (s.InsulationLayerThickness / 1000) / Conductivity));

				var mat6ASlabGreenGuideCorrected = slabs.Where(s => s.SlabType == SlabType.BaseSlab || s.SlabType == SlabType.Roof).Sum(s => (s.NetArea * (s.InsulationLayerThickness / 1000) / Conductivity) * s.InsulationLayerMaterialPoint);

				var mat6ATotalThermalResistance = mat6AWallThermalResistance + mat6ASlabThermalResistance;
				var mat6ATotalGreenGuideCorrected = mat6AWallGreenGuideCorrected + mat6ASlabGreenGuideCorrected;
				var insulationIndex = mat6ATotalGreenGuideCorrected / mat6ATotalThermalResistance;
				certification.Add("Mat 6a: Insulation", insulationIndex > 2 ? "1 Credit" : "0 Credit");
				documentCertficate.Add("Mat6a", insulationIndex > 2 ? "1 Credit" : "0 Credit");
			}
			else if (selectedMaterials.Contains("Mat6a"))
			{
				certification.Add("Mat 6a: Insulation", "No Credit");
				documentCertficate.Add("Mat6a", "No Credit");
			}

			if (mat6B == "1" && selectedMaterials.Contains("Mat6b"))
			{
				var mat6BrsmWallArea =
					externalWalls.Where(ew => ew.InsulationLayerThickness > 0 && ew.InsulationMaterialName.Contains("RSM")).Sum(ew => ew.NetSideArea);

				var mat6BrsmSlabArea =
					slabs.Where(
						s => !string.IsNullOrEmpty(s.InsulationMaterialName) && s.InsulationMaterialName.Contains("RSM"))
						.Where(s => s.SlabType == SlabType.BaseSlab || s.SlabType == SlabType.Roof)
						.Sum(s => s.NetArea);
				var mat6BrsmRatio = (mat6BrsmWallArea + mat6BrsmSlabArea) /
									(externalWalls.Where(ew => ew.InsulationLayerThickness > 0)
										.Sum(ew => ew.NetSideArea) +
									 slabs.Where(s => !string.IsNullOrEmpty(s.InsulationMaterialName))
										 .Where(s => s.SlabType == SlabType.BaseSlab || s.SlabType == SlabType.Roof)
										 .Sum(s => s.NetArea));

				certification.Add("Mat 6b: Responsibly Sourced Insulation", mat6BrsmRatio >= (decimal)0.8 ? "1 Credit" : "0 Credit");
				documentCertficate.Add("Mat6b", mat6BrsmRatio >= (decimal)0.8 ? "1 Credit" : "0 Credit");
			}
			else if (selectedMaterials.Contains("Mat6b"))
			{
				certification.Add("Mat 6b: Responsibly Sourced Insulation", "No Credit");
				documentCertficate.Add("Mat6b", "No Credit");
			}

			#endregion

			#region mat7

			if (selectedMaterials.Contains("Mat7"))
			{
				certification.Add("Mat 7: Designing for Robustness", (mat7 == "1") ? "1 Credit" : "0 Credit");
				documentCertficate.Add("Mat7", (mat7 == "1") ? "1 Credit" : "0 Credit");
			}

			#endregion

			return certification;
		}
		#region letgo
		private static string CalculateMatOneCertification(decimal mat1Point, string projectType, string buildingType)
		{
			const string newBuild = "New Build";
			const string office = "Office";
			const string retail = "Retail";
			const string majorRefurbishment = "Major Refurbishment";

			if (mat1Point < 2)
			{
				return "No Credit";
			}

			else if (mat1Point >= 2 && mat1Point < 4)
			{
				if (projectType.Contains(newBuild) || projectType.Contains(majorRefurbishment))
				{
					return "1 Credit";
				}
				return "1 Credit";
			}

			else if (mat1Point >= 4 && mat1Point < 5)
			{
				if (!projectType.Contains(newBuild) && !projectType.Contains(majorRefurbishment))
				{
					return "2 Credits";
				}

				if (buildingType.Contains(office) || buildingType.Contains(retail))
				{
					return "1 Credit";
				}
				return "2 Credits";
			}

			else if (mat1Point >= 5 && mat1Point < 8)
			{
				if (!projectType.Contains(newBuild) && !projectType.Contains(majorRefurbishment))
				{
					return "One additional exemplary credit";
				}

				if (buildingType.Contains(office) || buildingType.Contains(retail))
				{
					return "2 Credits";
				}
				return "One additional exemplary credit";
			}

			else if (mat1Point >= 8 && mat1Point < 10)
			{
				if (!projectType.Contains(newBuild) && !projectType.Contains(majorRefurbishment))
				{
					return "One additional exemplary credit";
				}

				if (buildingType.Contains(office) || buildingType.Contains(retail))
				{
					return "3 Credits";
				}
				return "One additional exemplary credit";
			}

			else if (mat1Point >= 10 && mat1Point < 12)
			{
				if (!projectType.Contains(newBuild) && !projectType.Contains(majorRefurbishment))
				{
					return "One additional exemplary credit";
				}

				if (buildingType.Contains(office) || buildingType.Contains(retail))
				{
					return "4 Credits";
				}
				return "One additional exemplary credit";
			}

			else
			{
				return "One additional exemplary credit";
			}
		}

		private List<Beam> RetrieveBeams()
		{
			return dataset.Tables["IfcBeam"].AsEnumerable().Select(CreateBeamFromRow).ToList();
		}

		private List<StairFlight> RetrieveStairs()
		{
			return dataset.Tables["IfcStairFlight"].AsEnumerable().Select(CreateStairFlightFromRow).ToList();
		}

		private List<Column> RetrieveColumns()
		{
			return dataset.Tables["IfcColumn"].AsEnumerable().Select(CreateColumnFromRow).ToList();
		}

		private List<Slab> RetrieveSlabs()
		{
			return dataset.Tables["IfcSlab"].AsEnumerable().Select(CreateSlabFromRow).ToList();
		}

		private List<Window> RetrieveWindows()
		{
			return dataset.Tables["IfcWindow"].AsEnumerable().Select(CreateWindowFromRow).ToList();
		}

		private List<Wall> RetrieveWallsByPosition(WallPositionType wallPositionType)
		{
			bool ifcWallExists = dataset.Tables.Contains("IfcWall");

			List<string> wallPropertySingleValueIds;
			if (wallPositionType == WallPositionType.External)
			{
				wallPropertySingleValueIds = RetrievePropertySingleValueIds("IsExternal", "1");

			}
			else if (wallPositionType == WallPositionType.Interior)
			{
				wallPropertySingleValueIds = RetrievePropertySingleValueIds("IsExternal", "0");
			}
			else
			{
				wallPropertySingleValueIds = RetrievePropertySingleValueIds("LoadBearing", "1");
			}

			var propertySetIds = new List<string>();
			var walls = new List<Wall>();

			if (wallPositionType == WallPositionType.Boundary)
			{
				propertySetIds = dataset.Tables["IfcPropertySet"].AsEnumerable()
					.Where(
						r =>
							r["Name"].ToString().Trim() == "Pset_WallCommon" &&
							r["HasProperties"].ToString().Contains("LoadBearing") &&
							!r["HasProperties"].ToString().Contains("IsExternal"))
					.Select(r => r[IdColumn].ToString()).ToList();
			}
			else
			{
				foreach (var wallPropertySingleValueId in wallPropertySingleValueIds)
				{
					var id = wallPropertySingleValueId;
					if (!dataset.Tables["IfcPropertySet"].AsEnumerable().Any(row => row["Name"].ToString().Trim() == "Pset_WallCommon" && row["HasProperties"].ToString().Contains(id)))
					{
						continue;
					}
					propertySetIds.Add(dataset.Tables["IfcPropertySet"].AsEnumerable().First(row => row["Name"].ToString().Trim() == "Pset_WallCommon" && row["HasProperties"].ToString().Contains(id))[IdColumn].ToString());
				}
			}


			if (!propertySetIds.Any())
			{
				return walls;
			}

			var wallIds = propertySetIds.Select(id => dataset.Tables["IfcWallStandardCase"].AsEnumerable().First(row => ((DataRow)row)["INV-IsDefinedBy"].ToString().Contains(id))[IdColumn].ToString()).ToList();

			if (!wallIds.Any())
			{
				return walls;
			}

			walls = wallIds.Select(externalWallId => dataset.Tables["IfcWallStandardCase"].AsEnumerable().First(r => r[IdColumn].ToString().Trim() == externalWallId)).Select(CreateWallFromRow).ToList();
			foreach (var wall in walls)
			{
				wall.WallPositionType = wallPositionType;
			}

			if (ifcWallExists)
			{
				var curvedWallIds = propertySetIds.Select(id => dataset.Tables["IfcWall"].AsEnumerable().First(row => ((DataRow)row)["INV-IsDefinedBy"].ToString().Contains(id))[IdColumn].ToString()).ToList();
				if (curvedWallIds.Any())
				{
					walls.AddRange(curvedWallIds.Select(externalWallId => dataset.Tables["IfcWall"].AsEnumerable().First(r => r[IdColumn].ToString().Trim() == externalWallId)).Select(CreateWallFromRow));
				}
			}


			return walls;
		}

		private List<Wall> RetrieveCurtainWalls()
		{
			var curtainWallRows = dataset.Tables["IfcCurtainWall"].AsEnumerable();
			var curtainWalls = curtainWallRows.Select(CreateWallFromRow).ToList();
			return curtainWalls;
		}

		private void RetrieveInsulationInformation(string ifcMaterialLayerSetUsagePart, out decimal insulationLayerThickness, out decimal insulationMaterialPoint, out string insulationMaterialName, out string insulationMaterialElementNumber)
		{
			var materialLayerIds =
				dataset.Tables["IfcMaterialLayer"].AsEnumerable().Where(r => r["Material"].ToString().ToLower().Contains("insulation") || r["Material"].ToString().ToLower().Contains("xps") || r["Material"].ToString().ToLower().Contains("eps")).Select(r => r[IdColumn].ToString()).ToList();
			var materialLayerMaterialLayerSetRelation = new Dictionary<string, string>();
			foreach (var materialLayerId in materialLayerIds)
			{
				var mlid = materialLayerId;
				var materialLayerSetId =
					dataset.Tables["IfcMaterialLayerSet"].AsEnumerable()
						.First(r => r["MaterialLayers"].ToString().Contains(mlid))[IdColumn].ToString();
				materialLayerMaterialLayerSetRelation.Add(mlid, materialLayerSetId);
			}

			var materialLayerSetMaterialLayerSetUsageRelation = new Dictionary<string, string>();
			foreach (var kvp in materialLayerMaterialLayerSetRelation)
			{
				var materialLayerSetUsageId =
					dataset.Tables["IfcMaterialLayerSetUsage"].AsEnumerable()
						.First(r => r["ForLayerSet"].ToString().Contains(kvp.Value))[IdColumn].ToString();
				if (materialLayerSetMaterialLayerSetUsageRelation.ContainsKey(kvp.Value))
				{
					continue;
				}
				materialLayerSetMaterialLayerSetUsageRelation.Add(kvp.Value, materialLayerSetUsageId);
			}

			var materialLayerSetUsageIds = new List<string>();
			foreach (var kvp in materialLayerSetMaterialLayerSetUsageRelation)
			{
				if (ifcMaterialLayerSetUsagePart.Contains(kvp.Value))
				{
					materialLayerSetUsageIds.Add(kvp.Value);
				}
			}

			insulationLayerThickness = 0;
			insulationMaterialName = string.Empty;
			insulationMaterialPoint = 0;
			insulationMaterialElementNumber = string.Empty;

			if (!materialLayerSetUsageIds.Any())
			{
				return;
			}

			var selectedMaterialLayerSetUsageId = materialLayerSetUsageIds.First();
			var selectedMaterialLayerSetId =
				materialLayerSetMaterialLayerSetUsageRelation.First(kvp => kvp.Value == selectedMaterialLayerSetUsageId)
					.Key;
			var selectedMaterialLayerIds =
				materialLayerMaterialLayerSetRelation.Where(kvp => kvp.Value == selectedMaterialLayerSetId)
					.AsEnumerable()
					.Select(kvp => kvp.Key)
					.ToList();


			foreach (var selectedMaterialLayerId in selectedMaterialLayerIds)
			{
				decimal layerThickness;
				var row = dataset.Tables["IfcMaterialLayer"].AsEnumerable().First(r => r[IdColumn].ToString() == selectedMaterialLayerId);
				if (decimal.TryParse(row["Column2"].ToString(), out layerThickness))
				{
					insulationLayerThickness += layerThickness;
					var materialName =
						dataset.Tables["IfcMaterial"].AsEnumerable()
							.First(r => row["Material"].ToString().Split(' ')[1] == r[IdColumn].ToString())["Name"]
							.ToString();
					insulationMaterialName += ((!string.IsNullOrEmpty(insulationMaterialName)) ? ", " : string.Empty) +
											  materialName;
					if (materialName.Contains("-"))
					{
						var elementNumberString = materialName.Trim().Split('-').Last().Trim();
						long elementNumber;
						if (long.TryParse(elementNumberString, out elementNumber))
						{
							var point = RetrieveMaterialPoint(elementNumberString, buildingType);
							insulationMaterialPoint += (layerThickness * point);
							insulationMaterialElementNumber = elementNumberString;
						}
					}
				}
			}
			if (insulationLayerThickness != 0)
			{
				insulationMaterialPoint = insulationMaterialPoint / insulationLayerThickness;
			}
		}

		private bool IsRsm(string propertySet)
		{
			var rsmPropertySingleValueRows =
				dataset.Tables["IfcPropertySingleValue"].AsEnumerable()
					.Where(r => r["Name"].ToString() == "ResponsiblySourced" && r["Column2"].ToString().Trim() == "1");
			if (!rsmPropertySingleValueRows.Any())
			{
				return false;
			}

			var rsmPropertySingleValueIds = rsmPropertySingleValueRows.Select(r => r[IdColumn].ToString());
			var propertySetIds = new List<string>();
			foreach (var rsmPropertySingleValueId in rsmPropertySingleValueIds)
			{
				var id = rsmPropertySingleValueId;
				propertySetIds.AddRange(dataset.Tables["IfcPropertySet"].AsEnumerable().Where(r => r["HasProperties"].ToString().Contains(id)).Select(r => r[IdColumn].ToString()));
			}

			return propertySetIds.Any(propertySet.Contains);
		}

		private bool IsFacadeReused(string propertySet)
		{
			var facadeReusedPropertySingleValueRows =
				dataset.Tables["IfcPropertySingleValue"].AsEnumerable()
					.Where(r => r["Name"].ToString() == "FacadeReused" && r["Column2"].ToString().Trim() == "1");
			if (!facadeReusedPropertySingleValueRows.Any())
			{
				return false;
			}

			var facadeReusedPropertySingleValueIds = facadeReusedPropertySingleValueRows.Select(r => r[IdColumn].ToString());
			var propertySetIds = new List<string>();
			foreach (var facadeReusedPropertySingleValueId in facadeReusedPropertySingleValueIds)
			{
				var id = facadeReusedPropertySingleValueId;
				propertySetIds.AddRange(dataset.Tables["IfcPropertySet"].AsEnumerable().Where(r => r["HasProperties"].ToString().Contains(id)).Select(r => r[IdColumn].ToString()));
			}

			return propertySetIds.Any(propertySet.Contains);
		}

		private bool IsStructureReused(string propertySet)
		{
			var structureReusedPropertySingleValueRows =
				dataset.Tables["IfcPropertySingleValue"].AsEnumerable()
					.Where(r => r["Name"].ToString() == "StructureReused" && r["Column2"].ToString().Trim() == "1");
			if (!structureReusedPropertySingleValueRows.Any())
			{
				return false;
			}

			var structureReusedPropertySingleValueIds = structureReusedPropertySingleValueRows.Select(r => r[IdColumn].ToString());
			var propertySetIds = new List<string>();
			foreach (var facadeReusedPropertySingleValueId in structureReusedPropertySingleValueIds)
			{
				var id = facadeReusedPropertySingleValueId;
				propertySetIds.AddRange(dataset.Tables["IfcPropertySet"].AsEnumerable().Where(r => r["HasProperties"].ToString().Contains(id)).Select(r => r[IdColumn].ToString()));
			}

			return propertySetIds.Any(propertySet.Contains);
		}

		private bool IsLoadBearing(string propertySet)
		{
			var structureReusedPropertySingleValueRows =
				dataset.Tables["IfcPropertySingleValue"].AsEnumerable()
					.Where(r => r["Name"].ToString() == "LoadBearing" && r["Column2"].ToString().Trim() == "1");
			if (!structureReusedPropertySingleValueRows.Any())
			{
				return false;
			}

			var structureReusedPropertySingleValueIds = structureReusedPropertySingleValueRows.Select(r => r[IdColumn].ToString());
			var propertySetIds = new List<string>();
			foreach (var facadeReusedPropertySingleValueId in structureReusedPropertySingleValueIds)
			{
				var id = facadeReusedPropertySingleValueId;
				propertySetIds.AddRange(dataset.Tables["IfcPropertySet"].AsEnumerable().Where(r => r["HasProperties"].ToString().Contains(id)).Select(r => r[IdColumn].ToString()));
			}

			return propertySetIds.Any(propertySet.Contains);
		}

		private Wall CreateWallFromRow(DataRow row)
		{
			var wall = new Wall();
			var isDefinedByCellContent = row["INV-IsDefinedBy"].ToString();
			var quantityPartOfIsDefinedByCellContent = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcElementQuantity"));
			var reverseString = string.Join(string.Empty, quantityPartOfIsDefinedByCellContent.Trim().Reverse());
			var quantityNumberReverse = reverseString.Substring(0, reverseString.IndexOf(" ", StringComparison.Ordinal));
			var quantityNumber = string.Join(string.Empty, quantityNumberReverse.Reverse());
			var elementQuantitiesInfo = dataset.Tables["IfcElementQuantity"].AsEnumerable().First(dr => dr[IdColumn].ToString() == quantityNumber && dr["Name"].ToString() == "BaseQuantities")["Quantities"].ToString();

			var propertySetPart = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet"));
			wall.Name = row["Name"].ToString();

			wall.Rsm = IsRsm(propertySetPart);
			wall.FacadeReuse = IsFacadeReused(propertySetPart);
			wall.StructureReuse = IsStructureReused(propertySetPart);
			wall.LoadBearing = IsLoadBearing(propertySetPart);

			if (wall.Rsm)
			{
				decimal rsmPoint;
				string rsmText;
				RetrieveRsmInformation(propertySetPart, out rsmText, out rsmPoint);
				wall.RsmPoint = rsmPoint;
				wall.RsmText = rsmText;

			}

			foreach (var quantityRow in dataset.Tables["IfcQuantityArea"].AsEnumerable().Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{
				switch (quantityRow["Name"].ToString())
				{
					case "GrossFootprintArea":
						decimal grossFootprintArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out grossFootprintArea))
						{
							wall.GrossFootprintArea = grossFootprintArea;
						}
						break;
					case "NetFootprintArea":
						decimal netFootprintArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netFootprintArea))
						{
							wall.NetFootprintArea = netFootprintArea;
						}
						break;
					case "GrossSideArea":
						decimal grossSideArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out grossSideArea))
						{
							wall.GrossSideArea = grossSideArea;
						}
						break;
					case "NetSideArea":
						decimal netSideArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netSideArea))
						{
							wall.NetSideArea = netSideArea;
						}
						break;
					default:
						break;

				}
			}

			foreach (
				var quantityRow in
					dataset.Tables["IfcQuantityVolume"].AsEnumerable()
						.Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{

				switch (quantityRow["Name"].ToString())
				{
					case "NetVolume":
						decimal netVolume;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netVolume))
						{
							wall.NetVolume = netVolume;
						}
						break;
					default:
						break;
				}

			}

			var hasAssociationsCell = row["INV-HasAssociations"].ToString();
			if (hasAssociationsCell.Contains("IfcClassificationReference"))
			{
				var ifcClassificationReferencePart =
				 hasAssociationsCell.Split('\n').First(s => s.Contains("IfcClassificationReference"));
				var ifcClassificationReferencePartReverse = new string(ifcClassificationReferencePart.Trim().Reverse().ToArray());
				var classificationNumberReverse = ifcClassificationReferencePartReverse.Substring(0,
					ifcClassificationReferencePartReverse.IndexOf(" ", StringComparison.Ordinal));
				var classificationNumberReference = new string(classificationNumberReverse.Reverse().ToArray());
				var itemReference =
					dataset.Tables["IfcClassificationReference"].AsEnumerable()
						.First(dr => dr[IdColumn].ToString() == classificationNumberReference)["ItemReference"].ToString();
				wall.ElementNumber = itemReference;
				wall.Point = RetrieveMaterialPoint(itemReference, buildingType);
			}
			if (hasAssociationsCell.Contains("IfcMaterialLayerSetUsage"))
			{
				var ifcMaterialLayerSetUsagePart =
				hasAssociationsCell.Split('\n').First(s => s.Contains("IfcMaterialLayerSetUsage"));

				decimal insulationLayerThickness;
				decimal insulationMaterialPoint;
				string insulationMaterialName;
				string insulationMaterialElementNumber;

				RetrieveInsulationInformation(ifcMaterialLayerSetUsagePart, out insulationLayerThickness,
					out insulationMaterialPoint, out insulationMaterialName, out insulationMaterialElementNumber);
				wall.InsulationLayerThickness = insulationLayerThickness;
				wall.InsulationLayerMaterialPoint = insulationMaterialPoint;
				wall.InsulationMaterialName = insulationMaterialName;
				wall.InsulationMaterialElementNumber = insulationMaterialElementNumber;
			}

			return wall;
		}

		private Window CreateWindowFromRow(DataRow row)
		{
			var window = new Window();
			window.Name = row["Name"].ToString();
			var isDefinedByCellContent = row["INV-IsDefinedBy"].ToString();
			var quantityPartOfIsDefinedByCellContent = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcElementQuantity"));
			var reverseString = new string(quantityPartOfIsDefinedByCellContent.Trim().Reverse().ToArray());
			var quantityNumberReverse = reverseString.Substring(0, reverseString.IndexOf(" ", StringComparison.Ordinal));
			var quantityNumber = new string(quantityNumberReverse.Reverse().ToArray());
			var elementQuantitiesInfo =
				dataset.Tables["IfcElementQuantity"].AsEnumerable()
					.First(dr => dr[IdColumn].ToString() == quantityNumber && dr["Name"].ToString() == "BaseQuantities")[
						"Quantities"].ToString();

			var propertySetPart = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet"));

			window.Rsm = IsRsm(isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet")));
			window.FacadeReuse = IsFacadeReused(propertySetPart);
			window.StructureReuse = IsStructureReused(propertySetPart);

			if (window.Rsm)
			{

				string tierText;
				decimal tierPoint;
				RetrieveRsmInformation(propertySetPart, out tierText, out tierPoint);
				window.RsmPoint = tierPoint;
				window.RsmText = tierText;
			}

			foreach (var quantityRow in dataset.Tables["IfcQuantityArea"].AsEnumerable().Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{
				switch (quantityRow["Name"].ToString())
				{
					case "GrossArea":
						decimal grossFootprintArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out grossFootprintArea))
						{
							window.GrossArea = grossFootprintArea;
						}
						break;
					case "Area":
						decimal netFootprintArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netFootprintArea))
						{
							window.NetArea = netFootprintArea;
						}
						break;
					default:
						break;

				}
			}

			var hasAssociationsCell = row["INV-HasAssociations"].ToString();
			var ifcClassificationReferencePart =
				hasAssociationsCell.Split('\n').First(s => s.Contains("IfcClassificationReference"));
			var ifcClassificationReferencePartReverse = new string(ifcClassificationReferencePart.Trim().Reverse().ToArray());
			var classificationNumberReverse = ifcClassificationReferencePartReverse.Substring(0,
				ifcClassificationReferencePartReverse.IndexOf(" ", StringComparison.Ordinal));
			var classificationNumberReference = new string(classificationNumberReverse.Reverse().ToArray());
			var itemReference =
				dataset.Tables["IfcClassificationReference"].AsEnumerable()
					.First(dr => dr[IdColumn].ToString() == classificationNumberReference)["ItemReference"].ToString();
			window.ElementNumber = itemReference;
			window.Point = RetrieveMaterialPoint(itemReference, buildingType);

			return window;
		}

		private void RetrieveRsmInformation(string propertySetPart, out string rsmText, out decimal rsmPoint)
		{
			var tierRowIds =
				dataset.Tables["IfcPropertyEnumeratedValue"].AsEnumerable()
					.Where(r => r["Name"].ToString() == "TierLevel").Select(r => r[IdColumn].ToString());


			rsmText = string.Empty;
			rsmPoint = 0;
			if (!tierRowIds.Any())
			{
				return;
			}

			var tierPropertySetRelations = new Dictionary<string, string>();

			foreach (var tierRowId in tierRowIds)
			{
				var trid = tierRowId;
				var propertySetId =
					dataset.Tables["IfcPropertySet"].AsEnumerable()
						.First(
							r =>
								r["HasProperties"].ToString().Contains("TierLevel") &&
								r["HasProperties"].ToString().Contains(trid))[IdColumn].ToString();
				tierPropertySetRelations.Add(trid, propertySetId);
			}

			foreach (var kvp in tierPropertySetRelations)
			{
				if (propertySetPart.Contains(kvp.Value))
				{
					var tierPropertySetRelation = kvp;
					var tierRow =
						dataset.Tables["IfcPropertyEnumeratedValue"].AsEnumerable()
							.First(r => r[IdColumn].ToString() == tierPropertySetRelation.Key);
					rsmText = tierRow["EnumerationValues"].ToString();
					if (rsmText.Contains("Excellent") || rsmText.Contains("Very Good"))
					{
						rsmPoint = 3;
					}
					else if (rsmText.Contains("Good") || rsmText.Contains("Good"))
					{
						rsmPoint = 2;
					}
					else if (rsmText.Contains("Certified EMS (for the Key Process and Supply Chain)") ||
							 rsmText.Contains("Verified (SmartWood) (for Timber only)") || rsmText.Contains("Certified EMS (for the Key Process) (for recycled materials only)"))
					{
						rsmPoint = (decimal)1.5;
					}

					else if (rsmText.Contains("Certified EMS (for key process stage)"))
					{
						rsmPoint = 1;
					}
					else
					{
						rsmPoint = 0;
					}
				}
			}
		}

		private Slab CreateSlabFromRow(DataRow row)
		{
			var slab = new Slab();
			switch (row["PredefinedType"].ToString())
			{
				case "BASESLAB":
					slab.SlabType = SlabType.BaseSlab;
					break;
				case "LANDING":
					slab.SlabType = SlabType.Landing;
					break;
				case "FLOOR":
					slab.SlabType = SlabType.Floor;
					break;
				case "ROOF":
					slab.SlabType = SlabType.Roof;
					break;
				default:
					slab.SlabType = SlabType.NotDefined;
					break;
			}
			slab.Name = row["Name"].ToString();
			var isDefinedByCellContent = row["INV-IsDefinedBy"].ToString();
			if (isDefinedByCellContent.Contains("IfcElementQuantity"))
			{
				var quantityPartOfIsDefinedByCellContent = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcElementQuantity"));
				var reverseString = new string(quantityPartOfIsDefinedByCellContent.Trim().Reverse().ToArray());
				var quantityNumberReverse = reverseString.Substring(0, reverseString.IndexOf(" ", StringComparison.Ordinal));
				var quantityNumber = new string(quantityNumberReverse.Reverse().ToArray());
				var elementQuantitiesInfo =
					dataset.Tables["IfcElementQuantity"].AsEnumerable()
						.First(dr => dr[IdColumn].ToString() == quantityNumber && dr["Name"].ToString() == "BaseQuantities")[
							"Quantities"].ToString();

				foreach (var quantityRow in dataset.Tables["IfcQuantityArea"].AsEnumerable().Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
				{
					switch (quantityRow["Name"].ToString())
					{
						case "GrossArea":
							decimal grossFootprintArea;
							if (decimal.TryParse(quantityRow["Column2"].ToString(), out grossFootprintArea))
							{
								slab.GrossArea = grossFootprintArea;
							}
							break;
						case "NetArea":
							decimal netFootprintArea;
							if (decimal.TryParse(quantityRow["Column2"].ToString(), out netFootprintArea))
							{
								slab.NetArea = netFootprintArea;
							}
							break;
						default:
							break;

					}
				}

				foreach (
				var quantityRow in
					dataset.Tables["IfcQuantityVolume"].AsEnumerable()
						.Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
				{

					switch (quantityRow["Name"].ToString())
					{
						case "NetVolume":
							decimal netVolume;
							if (decimal.TryParse(quantityRow["Column2"].ToString(), out netVolume))
							{
								slab.NetVolume = netVolume;
							}
							break;
						default:
							break;
					}

				}
			}

			var propertySetPart = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet"));

			slab.Rsm = IsRsm(propertySetPart);
			slab.FacadeReuse = IsFacadeReused(propertySetPart);
			slab.StructureReuse = IsStructureReused(propertySetPart);
			slab.LoadBearing = IsLoadBearing(propertySetPart);

			if (slab.Rsm)
			{
				string tierText;
				decimal tierPoint;
				RetrieveRsmInformation(propertySetPart, out tierText, out tierPoint);
				slab.RsmPoint = tierPoint;
				slab.RsmText = tierText;
			}

			var hasAssociationsCell = row["INV-HasAssociations"].ToString();
			if (hasAssociationsCell.Contains("IfcClassificationReference"))
			{
				var ifcClassificationReferencePart =
				hasAssociationsCell.Split('\n').First(s => s.Contains("IfcClassificationReference"));
				var ifcClassificationReferencePartReverse = new string(ifcClassificationReferencePart.Trim().Reverse().ToArray());
				var classificationNumberReverse = ifcClassificationReferencePartReverse.Substring(0,
					ifcClassificationReferencePartReverse.IndexOf(" ", StringComparison.Ordinal));
				var classificationNumberReference = new string(classificationNumberReverse.Reverse().ToArray());
				var itemReference =
					dataset.Tables["IfcClassificationReference"].AsEnumerable()
						.First(dr => dr[IdColumn].ToString() == classificationNumberReference)["ItemReference"].ToString();
				slab.ElementNumber = itemReference;
				slab.Point = RetrieveMaterialPoint(itemReference, buildingType);
			}

			if (hasAssociationsCell.Contains("IfcMaterialLayerSetUsage"))
			{
				var ifcMaterialLayerSetUsagePart =
				hasAssociationsCell.Split('\n').First(s => s.Contains("IfcMaterialLayerSetUsage"));

				decimal insulationLayerThickness;
				decimal insulationMaterialPoint;
				string insulationMaterialName;
				string insulationMaterialElementNumber;

				RetrieveInsulationInformation(ifcMaterialLayerSetUsagePart, out insulationLayerThickness,
					out insulationMaterialPoint, out insulationMaterialName, out insulationMaterialElementNumber);
				slab.InsulationLayerThickness = insulationLayerThickness;
				slab.InsulationLayerMaterialPoint = insulationMaterialPoint;
				slab.InsulationMaterialName = insulationMaterialName;
				slab.InsulationMaterialElementNumber = insulationMaterialElementNumber;
			}
			return slab;
		}

		private Column CreateColumnFromRow(DataRow row)
		{
			var column = new Column();
			column.Name = row["Name"].ToString();
			var isDefinedByCellContent = row["INV-IsDefinedBy"].ToString();
			var quantityPartOfIsDefinedByCellContent = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcElementQuantity"));
			var reverseString = new string(quantityPartOfIsDefinedByCellContent.Trim().Reverse().ToArray());
			var quantityNumberReverse = reverseString.Substring(0, reverseString.IndexOf(" ", StringComparison.Ordinal));
			var quantityNumber = new string(quantityNumberReverse.Reverse().ToArray());
			var elementQuantitiesInfo =
				dataset.Tables["IfcElementQuantity"].AsEnumerable()
					.First(dr => dr[IdColumn].ToString() == quantityNumber && dr["Name"].ToString() == "BaseQuantities")[
						"Quantities"].ToString();
			foreach (var quantityRow in dataset.Tables["IfcQuantityArea"].AsEnumerable().Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{
				switch (quantityRow["Name"].ToString())
				{
					case "CrossSectionArea":
						decimal crossSectionArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out crossSectionArea))
						{
							column.CrossSectionArea = crossSectionArea;
						}
						break;
					case "OuterSurfaceArea":
						decimal outerSurfaceArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out outerSurfaceArea))
						{
							column.OuterSurfaceArea = outerSurfaceArea;
						}
						break;
					case "TotalSurfaceArea":
						decimal totalSurfaceArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out totalSurfaceArea))
						{
							column.TotalSurfaceArea = totalSurfaceArea;
						}
						break;
					case "GrossSurfaceArea":
						decimal grossSurfaceArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out grossSurfaceArea))
						{
							column.GrossSurfaceArea = grossSurfaceArea;
						}
						break;
					default:
						break;
				}
			}
			foreach (
				var quantityRow in
					dataset.Tables["IfcQuantityVolume"].AsEnumerable()
						.Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{

				switch (quantityRow["Name"].ToString())
				{
					case "NetVolume":
						decimal netVolume;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netVolume))
						{
							column.NetVolume = netVolume;
						}
						break;
					default:
						break;
				}

			}

			var propertySetPart = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet"));

			column.Rsm = IsRsm(propertySetPart);
			column.FacadeReuse = IsFacadeReused(propertySetPart);
			column.StructureReuse = IsStructureReused(propertySetPart);

			if (column.Rsm)
			{
				string tierText;
				decimal tierPoint;
				RetrieveRsmInformation(propertySetPart, out tierText, out tierPoint);
				column.RsmPoint = tierPoint;
				column.RsmText = tierText;
			}


			var hasAssociationsCell = row["INV-HasAssociations"].ToString();
			if (hasAssociationsCell.Contains("IfcClassificationReference"))
			{
				var ifcClassificationReferencePart = hasAssociationsCell.Split('\n').First(s => s.Contains("IfcClassificationReference"));
				var ifcClassificationReferencePartReverse = new string(ifcClassificationReferencePart.Trim().Reverse().ToArray());
				var classificationNumberReverse = ifcClassificationReferencePartReverse.Substring(0,
					ifcClassificationReferencePartReverse.IndexOf(" ", StringComparison.Ordinal));
				var classificationNumberReference = new string(classificationNumberReverse.Reverse().ToArray());
				var itemReference =
					dataset.Tables["IfcClassificationReference"].AsEnumerable()
						.First(dr => dr["ID"].ToString() == classificationNumberReference)["ItemReference"].ToString();
				column.ElementNumber = itemReference;
				column.Point = RetrieveMaterialPoint(itemReference, buildingType);
			}

			return column;
		}

		private StairFlight CreateStairFlightFromRow(DataRow row)
		{
			var isDefinedByCellContent = row["INV-IsDefinedBy"].ToString();
			var propertySetPart = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet"));
			var stairFlight = new StairFlight
			{
				NumberOfRiser = int.Parse(row["NumberOfRiser"].ToString()),
				NumberOfTreads = int.Parse(row["NumberOfTreads"].ToString()),
				RiserHeight = decimal.Parse(row["RiserHeight"].ToString()),
				TreadLength = decimal.Parse(row["TreadLength"].ToString()),
				TreadDistance = (decimal)0.90,
				FacadeReuse = IsFacadeReused(propertySetPart),
				Rsm = IsRsm(propertySetPart),
				StructureReuse = IsStructureReused(propertySetPart),
				Name = row["Name"].ToString(),
			};

			if (stairFlight.Rsm)
			{
				string tierText;
				decimal tierPoint;
				RetrieveRsmInformation(propertySetPart, out tierText, out tierPoint);
				stairFlight.RsmPoint = tierPoint;
				stairFlight.RsmText = tierText;
			}

			return stairFlight;
		}

		private Beam CreateBeamFromRow(DataRow row)
		{
			var beam = new Beam();
			beam.Name = row["Name"].ToString();
			var isDefinedByCellContent = row["INV-IsDefinedBy"].ToString();
			var quantityPartOfIsDefinedByCellContent = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcElementQuantity"));
			var reverseString = new string(quantityPartOfIsDefinedByCellContent.Trim().Reverse().ToArray());
			var quantityNumberReverse = reverseString.Substring(0, reverseString.IndexOf(" ", StringComparison.Ordinal));
			var quantityNumber = new string(quantityNumberReverse.Reverse().ToArray());
			var elementQuantitiesInfo =
				dataset.Tables["IfcElementQuantity"].AsEnumerable()
					.First(dr => dr[IdColumn].ToString() == quantityNumber && dr["Name"].ToString() == "BaseQuantities")[
						"Quantities"].ToString();
			foreach (var quantityRow in dataset.Tables["IfcQuantityArea"].AsEnumerable().Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{
				switch (quantityRow["Name"].ToString())
				{
					case "CrossSectionArea":
						decimal crossSectionArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out crossSectionArea))
						{
							beam.CrossSectionArea = crossSectionArea;
						}
						break;
					case "OuterSurfaceArea":
						decimal outerSurfaceArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out outerSurfaceArea))
						{
							beam.OuterSurfaceArea = outerSurfaceArea;
						}
						break;
					case "TotalSurfaceArea":
						decimal totalSurfaceArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out totalSurfaceArea))
						{
							beam.TotalSurfaceArea = totalSurfaceArea;
						}
						break;
					case "GrossSurfaceArea":
						decimal grossSurfaceArea;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out grossSurfaceArea))
						{
							beam.GrossSurfaceArea = grossSurfaceArea;
						}
						break;
					case "NetSurfaceAreaExtrudedSides":
						decimal netSurfaceAreaExtrudedSides;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netSurfaceAreaExtrudedSides))
						{
							beam.NetSurfaceAreaExtrudedSides = netSurfaceAreaExtrudedSides;
						}
						break;
					default:
						break;
				}
			}

			foreach (
				var quantityRow in
					dataset.Tables["IfcQuantityVolume"].AsEnumerable()
						.Where(dr => elementQuantitiesInfo.Contains(dr[IdColumn].ToString())))
			{

				switch (quantityRow["Name"].ToString())
				{
					case "NetVolume":
						decimal netVolume;
						if (decimal.TryParse(quantityRow["Column2"].ToString(), out netVolume))
						{
							beam.NetVolume = netVolume;
						}
						break;
					default:
						break;
				}

			}

			var propertySetPart = isDefinedByCellContent.Split('\n').First(s => s.Contains("IfcPropertySet"));

			beam.Rsm = IsRsm(propertySetPart);
			beam.FacadeReuse = IsFacadeReused(propertySetPart);
			beam.StructureReuse = IsStructureReused(propertySetPart);

			if (beam.Rsm)
			{
				string tierText;
				decimal tierPoint;
				RetrieveRsmInformation(propertySetPart, out tierText, out tierPoint);
				beam.RsmPoint = tierPoint;
				beam.RsmText = tierText;
			}


			var hasAssociationsCell = row["INV-HasAssociations"].ToString();
			if (hasAssociationsCell.Contains("IfcClassificationReference"))
			{
				var ifcClassificationReferencePart = hasAssociationsCell.Split('\n').First(s => s.Contains("IfcClassificationReference"));
				var ifcClassificationReferencePartReverse = new string(ifcClassificationReferencePart.Trim().Reverse().ToArray());
				var classificationNumberReverse = ifcClassificationReferencePartReverse.Substring(0,
					ifcClassificationReferencePartReverse.IndexOf(" ", StringComparison.Ordinal));
				var classificationNumberReference = new string(classificationNumberReverse.Reverse().ToArray());
				var itemReference =
					dataset.Tables["IfcClassificationReference"].AsEnumerable()
						.First(dr => dr["ID"].ToString() == classificationNumberReference)["ItemReference"].ToString();
				beam.ElementNumber = itemReference;
				beam.Point = RetrieveMaterialPoint(itemReference, buildingType);
			}

			return beam;
		}

		private decimal RetrieveMaterialPoint(string itemReference, string buildingType)
		{
			var rows = materialTable.AsEnumerable().Where(r => r["Element Number"].ToString() == itemReference);
			DataRow row;

			if (rows.Count() < 1)
			{
				return 0;
			}
			else if (rows.Count() == 1)
			{
				row = rows.First();
			}
			else
			{
				row = rows.FirstOrDefault(r => r["Building Type"].ToString().ToLower().Contains(buildingType.ToLower()));
				if (row == null)
				{
					row = rows.First();
				}
			}

			if (row == null)
			{
				return 0;
			}

			var summaryRating = row["Summary Rating"].ToString();
			switch (summaryRating)
			{
				case "A+":
					return 3;
					break;
				case "A":
					return 2;
					break;
				case "B":
					return 1;
					break;
				case "C":
					return (decimal)0.5;
					break;
				case "D":
					return (decimal)0.25;
					break;
				case "E":
					return 0;
					break;
				default:
					return 0;
					break;
			}
		}

		private List<string> RetrievePropertySingleValueIds(string name, string value)
		{

			if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
			{
				var rows =
					dataset.Tables["IfcPropertySingleValue"].AsEnumerable()
						.Where(r => r["Name"].ToString() == name && r["Column2"].ToString() == value);
				if (rows.Any())
				{
					return rows.Select(r => r[IdColumn].ToString()).ToList();
				}
				return new List<string>();
			}
			return new List<string>();

			//var query = string.Empty;
			//if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(value))
			//{
			//    throw new Exception("You should enter name or value");
			//}

			//if (!string.IsNullOrEmpty(name))
			//{
			//    query = string.Format("Name='{0}'", name);
			//}
			//if (!string.IsNullOrEmpty(value))
			//{
			//    query += (!string.IsNullOrEmpty(query) ? " AND " : string.Empty) +
			//             string.Format("Column2='{0}'", value);
			//}

			//var rows = dataset.Tables["IfcPropertySingleValue"].Select(query);
			//if (rows.Length < 0)
			//{
			//    return new List<string>();
			//}
			//return rows.Select(r => ((DataRow)r)[IdColumn].ToString()).ToList();
		}

		private string RetrievePropertyNominalValueByName(string name)
		{
			var rows = dataset.Tables["IfcPropertySingleValue"].AsEnumerable().Where(r => r["Name"].ToString().Trim() == name).ToList();
			if (rows.Count < 1)
			{
				throw new Exception("Property name '" + name + "' not found");
			}
			var row = rows.First();
			return row["Column2"].ToString();
		}

		private string RetrieveEnumeratedValue(string name)
		{
			var rows = dataset.Tables["IfcPropertyEnumeratedValue"].AsEnumerable().Where(r => r["Name"].ToString().Trim() == name);
			if (!rows.Any())
			{
				throw new Exception("Building type not specified");
			}
			var row = rows.First();
			return row["EnumerationValues"].ToString();
		}

		private void ProcessOnExited(object sender, EventArgs eventArgs)
		{
			if (process.ExitCode != -1)
			{

				MainProcess();
			}
		}

		private void MainProcess()
		{
			this.Invoke((MethodInvoker)delegate { IFAProgress.Visible = false; });
			this.Invoke((MethodInvoker)delegate { CancelProcess.Enabled = false; });
			ProcessXlsxFile();
			var projectInformation = RetrieveProjectInformation();
			this.projectType = RetrieveEnumeratedValue("ProjectType").Replace("{", string.Empty).Replace("}", string.Empty);
			projectPhase = projectInformation["Phase"];
			projectName = projectInformation["LongName"];
			buildingType = RetrieveEnumeratedValue("BuildingType");

			var certification = CalculateCertification(projectType, buildingType);

			var resultText =
				string.Format("Project Name: {0}{1}Project Type: {2}{1}Building Type: {3}{1}Project Phase: {4}{1}",
					projectInformation["LongName"], Environment.NewLine, projectType, buildingType, projectInformation["Phase"]);

			//var resultText = string.Format("Project Type: {0}{1}Building Type :{2}{1}", projectType, Environment.NewLine, buildingType);

			resultText = certification.Aggregate(resultText,
				(current, certificate) =>
					current + string.Format("{2}{0}: {1}", certificate.Key, certificate.Value, Environment.NewLine));
			this.Invoke((MethodInvoker)delegate { Results.Text = resultText; });
			this.Invoke((MethodInvoker)delegate { Report.Visible = true; });
		}

		private Dictionary<string, string> RetrieveProjectInformation()
		{
			var row = dataset.Tables["IfcProject"].Rows[0];
			var dict = new Dictionary<string, string>();
			dict.Add("Name", row["Name"].ToString());
			dict.Add("LongName", row["LongName"].ToString());
			dict.Add("Phase", row["Phase"].ToString());
			return dict;
		}

		private void ProcessXlsxFile()
		{
			var ifcFileName = IFCFileUrl.Text;
			var xlsxFileName = ifcFileName.Replace(".ifc", "_ifc") + ".xlsx";
			var connectionString = string.Format(ConfigurationManager.ConnectionStrings["ExcelConnectionString"].ConnectionString, xlsxFileName);
			var dummyDs = new DataSet();
			using (var connection = new OleDbConnection(connectionString))
			{
				connection.Open();
				var sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
				if (sheets == null)
				{
					return;
				}

				const string query = "SELECT * FROM [{0}]";
				foreach (DataRow sheet in sheets.Rows)
				{
					var adapter = new OleDbDataAdapter(string.Format(query, sheet["TABLE_NAME"]), connection);
					var table = new DataTable(sheet["TABLE_NAME"].ToString());
					adapter.Fill(table);
					dummyDs.Tables.Add(table);
				}
				dataset = new DataSet();
				foreach (DataTable datatable in dummyDs.Tables)
				{
					var rowStartIndex = 3;
					if (datatable.TableName == "Header$")
					{
						continue;
					}
					if (datatable.TableName == "Summary$")
					{
						rowStartIndex = 8;
					}
					var table = new DataTable(datatable.TableName.Replace("$", string.Empty));
					var headerRow = datatable.Rows[rowStartIndex - 1];
					foreach (var columnName in headerRow.ItemArray)
					{
						if (table.Columns.Contains(columnName.ToString()))
						{
							table.Columns.Add(columnName + "_" + Guid.NewGuid());
						}
						else
						{
							table.Columns.Add(columnName.ToString());
						}
					}

					for (var rowIndex = rowStartIndex; rowIndex < datatable.Rows.Count; rowIndex++)
					{
						var rowItems = datatable.Rows[rowIndex].ItemArray.ToArray();
						table.Rows.Add(rowItems);
					}

					dataset.Tables.Add(table);
				}
			}

			using (var connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["GMDB"].ConnectionString))
			{
				connection.Open();
				var sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
				const string query = "SELECT * FROM [MATERIALS$]";
				var adapter = new OleDbDataAdapter(query, connection);
				materialTable = new DataTable("MATERIALS");
				adapter.Fill(materialTable);
			}
		}

		private void CancelProcessClicked(object sender, EventArgs e)
		{
			process.Kill();
			Results.Text = "";
			IFAProgress.Visible = false;
			CancelProcess.Enabled = false;
		}

		private List<string> RetrieveSelectedMaterials()
		{
			var selectedItems = MaterialList.CheckedItems;
			var selectedMaterials = new List<string>();
			foreach (var selectedItem in selectedItems)
			{
				var type = selectedItem.GetType();
				var property = type.GetProperty("Value");
				selectedMaterials.Add(property.GetValue(selectedItem, null).ToString());
			}

			return selectedMaterials;
		}

		private void SelectAllClicked(object sender, EventArgs e)
		{

			for (var index = 0; index < MaterialList.Items.Count; index++)
			{
				MaterialList.SetItemChecked(index, true);
			}
		}

		private void DeselectAllClicked(object sender, EventArgs e)
		{
			for (var index = 0; index < MaterialList.Items.Count; index++)
			{
				MaterialList.SetItemCheckState(index, CheckState.Unchecked);
			}
		}
		#endregion
		private void CreatePdf(string pdfFilePath)
		{
			var pdfDoc = new pdfDocument("BREEAM Materials Report", "");

			var coverPage = pdfDoc.addPage(predefinedPageSize.csA4Page);

			coverPage.addText("BREEAM Materials", 40, coverPage.height - 140, pdfDoc.getFontReference(predefinedFont.csHelveticaBold), 20);
			coverPage.addText(string.Format("Project Name: {0}", projectName), 40, coverPage.height - 165, pdfDoc.getFontReference(predefinedFont.csHelvetica), 10);
			coverPage.addText(string.Format("Project Type: {0}", projectType), 40, coverPage.height - 180, pdfDoc.getFontReference(predefinedFont.csHelvetica), 10);
			coverPage.addText(string.Format("Building Type: {0}", buildingType), 40, coverPage.height - 195, pdfDoc.getFontReference(predefinedFont.csHelvetica), 10);
			coverPage.addText(string.Format("Project Phase: {0}", projectPhase), 40, coverPage.height - 210, pdfDoc.getFontReference(predefinedFont.csHelvetica), 10);

			if (documentCertficate.ContainsKey("Mat1"))
			{
				var mat1Items = new Dictionary<string, List<Dictionary<string, object>>>();

				var gr = externalWalls.GroupBy(ew => ew.ElementNumber);

				var mat1ExternalWallGroup =
					gr
					 .Select(
						 ew =>
							 new Dictionary<string, object>()
                            {
                                {"Name", ew.First().Name},
                                {"Element Number", ew.Key},
                                {"Area", ew.Sum(w => w.NetSideArea).ToString("N")},
                                {"Green Guide Rating", ew.First().Point}
                            });
				if (mat1ExternalWallGroup.Any())
				{
					mat1Items.Add("External Wall", mat1ExternalWallGroup.ToList());
				}

				var mat1CurtainWallGroup = curtainWalls.GroupBy(ew => ew.ElementNumber)
						.Select(
							ew =>
								new Dictionary<string, object>()
                            {
                                {"Name", ew.First().Name},
                                {"Element Number", ew.Key},
                                {"Area", ew.Sum(w => w.GrossSideArea).ToString("N")},
                                {"Green Guide Rating", ew.First().Point}
                            });
				if (mat1CurtainWallGroup.Any())
				{
					mat1Items.Add("Curtain Wall", mat1CurtainWallGroup.ToList());
				}

				var mat1WindowGroup = windows.GroupBy(ew => ew.ElementNumber)
						.Select(
							w =>
								new Dictionary<string, object>()
                            {
                                {"Name", w.First().Name},
                                {"Element Number", w.Key},
                                {"Area", w.Sum(wi => wi.NetArea).ToString("N")},
                                {"Green Guide Rating", w.First().Point}
                            });
				if (mat1WindowGroup.Any())
				{
					mat1Items.Add("Window", mat1WindowGroup.ToList());
				}

				var mat1FloorSlabsGroup = slabs.Where(s => s.SlabType == SlabType.Floor).GroupBy(s => s.ElementNumber)
						.Select(
							s =>
								new Dictionary<string, object>()
                            {
                                {"Name", s.First().Name},
                                {"Element Number", s.Key},
                                {"Area", s.Sum(sl => sl.NetArea).ToString("N")},
                                {"Green Guide Rating", s.First().Point}
                            });
				if (mat1FloorSlabsGroup.Any())
				{
					mat1Items.Add("Upper Floor", mat1FloorSlabsGroup.ToList());
				}

				var mat1RoofSlabGroup = slabs.Where(s => s.SlabType == SlabType.Roof).GroupBy(s => s.ElementNumber)
						.Select(
							s =>
								new Dictionary<string, object>()
                            {
                                {"Name", s.First().Name},
                                {"Element Number", s.Key},
                                {"Area", s.Sum(sl => sl.NetArea).ToString("N")},
                                {"Green Guide Rating", s.First().Point}
                            });
				if (mat1RoofSlabGroup.Any())
				{
					mat1Items.Add("Roof", mat1RoofSlabGroup.ToList());
				}
				if (!mat1Items.Any())
				{
					mat1Items.Add("Material Specification", new List<Dictionary<string, object>>());
				}
				AddTable("Mat 1", mat1Items, ref pdfDoc);
			}

			if (documentCertficate.ContainsKey("Mat2a"))
			{
				var mat2aItems = new Dictionary<string, List<Dictionary<string, object>>>();
				mat2aItems.Add("Natural Boundary", new List<Dictionary<string, object>>());
				AddTable("Mat 2a", mat2aItems, ref pdfDoc);
			}
			if (documentCertficate.ContainsKey("Mat2b"))
			{
				var mat2BItems = new Dictionary<string, List<Dictionary<string, object>>>();

				var mat2BBoundaryWallGroup = boundaryWalls.GroupBy(bw => bw.ElementNumber)
					.Select(
						bw =>
							new Dictionary<string, object>()
                            {
                                {"Name", bw.First().Name},
                                {"Element Number", bw.Key},
                                {"Area", bw.Sum(w => w.NetSideArea).ToString("N")},
                                {"Green Guide Rating", bw.First().Point}
                            });
				if (mat2BBoundaryWallGroup.Any())
				{
					mat2BItems.Add("Boundary Protection", mat2BBoundaryWallGroup.ToList());
				}

				var mat2BLandingSlabGroup = slabs.Where(s => s.SlabType == SlabType.Landing).GroupBy(s => s.ElementNumber)
						.Select(
							s =>
								new Dictionary<string, object>()
                            {
                                {"Name", s.First().Name},
                                {"ElementNumber", s.Key},
                                {"Area", s.Sum(sl => sl.NetArea).ToString("N")},
                                {"Green Guide Rating", s.First().Point}
                            });
				if (mat2BLandingSlabGroup.Any())
				{
					mat2BItems.Add("Hard Landscaping", mat2BLandingSlabGroup.ToList());
				}
				if (!mat2BItems.Any())
				{
					mat2BItems.Add("Hard Landscaping and Boundary Protection", new List<Dictionary<string, object>>());
				}
				AddTable("Mat 2b", mat2BItems, ref pdfDoc);
			}

			if (documentCertficate.ContainsKey("Mat3"))
			{

				var mat3Items = new Dictionary<string, List<Dictionary<string, object>>>();

				if (documentCertficate["Mat3"] == "No Credit")
				{
					mat3Items.Add("Façade Reuse", new List<Dictionary<string, object>>());
				}
				else
				{
					var mat3ExternalWallGroup = externalWalls.Where(ew => ew.FacadeReuse).GroupBy(ew => ew.ElementNumber)
					.Select(
						ew =>
							new Dictionary<string, object>()
                            {
                                {"Name", ew.First().Name},
                                {"ElementNumber", ew.Key},
                                {"Area", ew.Sum(w => w.NetSideArea).ToString("N")},
                                {"Percentage", (ew.Sum(w => w.NetSideArea)/externalWalls.Sum(w => w.NetSideArea)).ToString("P")}
                            });
					if (mat3ExternalWallGroup.Any())
					{
						mat3Items.Add("Reused External Wall", mat3ExternalWallGroup.ToList());
					}

					var mat3CurtainWallGroup = curtainWalls.Where(cw => cw.FacadeReuse).GroupBy(ew => ew.ElementNumber)
							.Select(
								cw =>
									new Dictionary<string, object>()
                            {
                                {"Name", cw.First().Name},
                                {"ElementNumber", cw.Key},
                                {"Area", cw.Sum(w => w.GrossSideArea).ToString("N")},
                                {"Percentage", (cw.Sum(w => w.GrossSideArea)/curtainWalls.Sum(w => w.GrossSideArea)).ToString("P")}
                            });
					if (mat3CurtainWallGroup.Any())
					{
						mat3Items.Add("Reused Curtain Wall", mat3CurtainWallGroup.ToList());
					}

					var mat3WindowGroup = windows.Where(w => w.FacadeReuse).GroupBy(w => w.ElementNumber)
							.Select(
								w =>
									new Dictionary<string, object>()
                            {
                                {"Name", w.First().Name},
                                {"ElementNumber", w.Key},
                                {"Area", w.Sum(wi => wi.NetArea).ToString("N")},
                                {"Percentage", (w.Sum(wi => wi.NetArea)/windows.Sum(wi => wi.NetArea)).ToString("P")}
                            });

					if (mat3WindowGroup.Any())
					{
						mat3Items.Add("Reused Window", mat3WindowGroup.ToList());
					}
					if (!mat3Items.Any())
					{
						mat3Items.Add("Façade Reuse", new List<Dictionary<string, object>>());
					}
				}
				
				AddTable("Mat 3", mat3Items, ref pdfDoc);

			}

			if (documentCertficate.ContainsKey("Mat4"))
			{
				var mat4Items = new Dictionary<string, List<Dictionary<string, object>>>();
				if (documentCertficate["Mat4"] == "No Credit")
				{
					mat4Items.Add("Structure Reuse", new List<Dictionary<string, object>>());
				}
				else
				{
					var mat4ExternalWallGroup = externalWalls.Where(ew => ew.LoadBearing && ew.StructureReuse).GroupBy(ew => ew.ElementNumber)
						.Select(
							ew =>
								new Dictionary<string, object>()
                            {
                                {"Name", ew.First().Name},
                                {"ElementNumber", ew.Key},
                                {"Volume", ew.Sum(w => w.NetVolume).ToString("N")},
                                {"Percentage", (ew.Sum(w => w.NetVolume)/externalWalls.Where(w => w.LoadBearing).Sum(w => w.NetVolume)).ToString("P")}
                            });
					if (mat4ExternalWallGroup.Any())
					{
						mat4Items.Add("Reused External Wall", mat4ExternalWallGroup.ToList());
					}

					var mat4InteriorWallGroup = interiorWalls.Where(iw => iw.LoadBearing && iw.StructureReuse).GroupBy(iw => iw.ElementNumber)
							.Select(
								cw =>
									new Dictionary<string, object>()
                            {
                                {"Name", cw.First().Name},
                                {"ElementNumber", cw.Key},
                                {"Volume", cw.Sum(w => w.NetVolume).ToString("N")},
                                {"Percentage", (cw.Sum(w => w.NetVolume)/interiorWalls.Where(w => w.LoadBearing).Sum(w => w.NetVolume)).ToString("P")}
                            });
					if (mat4InteriorWallGroup.Any())
					{
						mat4Items.Add("Reused Interior Wall", mat4InteriorWallGroup.ToList());
					}

					var mat4SlabsGroup = slabs.Where(s => s.SlabType == SlabType.Floor && s.LoadBearing && s.StructureReuse).GroupBy(s => s.ElementNumber)
							.Select(
								s =>
									new Dictionary<string, object>()
                            {
                                {"Name", s.First().Name},
                                {"ElementNumber", s.Key},
                                {"Volume", s.Sum(w => w.NetVolume).ToString("N")},
                                {"Percentage", (s.Sum(w => w.NetVolume)/slabs.Where(w => w.SlabType == SlabType.Floor && w.LoadBearing).Sum(w => w.NetVolume)).ToString("P")}
                            });

					if (mat4SlabsGroup.Any())
					{
						mat4Items.Add("Reused Slab", mat4SlabsGroup.ToList());
					}

					var mat4ColumnGroup = columns.Where(c => c.StructureReuse).GroupBy(c => c.ElementNumber)
							.Select(
								c =>
									new Dictionary<string, object>()
                            {
                                {"Name", c.First().Name},
                                {"ElementNumber", c.Key},
                                {"Volume", c.Sum(w => w.NetVolume).ToString("N")},
                                {"Percentage", (c.Sum(co => co.NetVolume)/columns.Sum(co => co.NetVolume)).ToString("P")}
                            });

					if (mat4ColumnGroup.Any())
					{
						mat4Items.Add("Reused Column", mat4ColumnGroup.ToList());
					}

					var mat4BeamGroup = beams.Where(b => b.StructureReuse).GroupBy(b => b.ElementNumber)
							.Select(
								b =>
									new Dictionary<string, object>()
                            {
                                {"Name", b.First().Name},
                                {"ElementNumber", b.Key},
                                {"Volume", b.Sum(w => w.NetVolume).ToString("N")},
                                {"Percentage", (b.Sum(be => be.NetVolume)/columns.Sum(be => be.NetVolume)).ToString("P")}
                            });
					if (mat4BeamGroup.Any())
					{
						mat4Items.Add("Reused Beam", mat4BeamGroup.ToList());
					}
					if (!mat4Items.Any())
					{
						mat4Items.Add("Structure Reuse", new List<Dictionary<string, object>>());
					}
				}
				AddTable("Mat 4", mat4Items, ref pdfDoc);
			}

			if (documentCertficate.ContainsKey("Mat5"))
			{
				var mat5Items = new Dictionary<string, List<Dictionary<string, object>>>();
				var mat5ExternalWallGroup = externalWalls.Where(ew => ew.Rsm).GroupBy(ew => ew.ElementNumber)
					.Select(
						ew =>
							new Dictionary<string, object>()
                            {
                                {"Name", ew.First().Name},
                                {"Element Number", ew.Key},
                                {"Area", ew.Sum(w => w.NetSideArea).ToString("N")},
                                {"Percentage", (ew.Sum(w => w.NetSideArea)/externalWalls.Sum(w => w.NetSideArea)).ToString("P")},
                                {"Tier Level", ew.First().RsmText}
                            });

				if (mat5ExternalWallGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Walls", mat5ExternalWallGroup.ToList());
				}

				var mat5InteriorWallGroup = interiorWalls.Where(iw => iw.Rsm).GroupBy(iw => iw.ElementNumber)
						.Select(
							iw =>
								new Dictionary<string, object>()
                            {
                                {"Name", iw.First().Name},
                                {"Element Number", iw.Key},
                                {"Area", iw.Sum(w => w.GrossSideArea).ToString("N")},
                                {"Percentage", (iw.Sum(w => w.GrossSideArea)/interiorWalls.Sum(w => w.GrossSideArea)).ToString("P")},
                                {"Tier Level", iw.First().RsmText}
                            });
				if (mat5InteriorWallGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Walls", mat5InteriorWallGroup.ToList());
				}

				var mat5RoofGroup =
					slabs.Where(s => s.Rsm && s.SlabType == SlabType.Roof)
						.GroupBy(s => s.ElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().Name},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        {"Percentage", (s.Sum(w => w.NetArea)/slabs.Where(sl => sl.SlabType == SlabType.Roof).Sum(w => w.NetArea)).ToString("P")},
                        {"Tier Level", s.First().RsmText}
                    });
				if (mat5RoofGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Roof", mat5RoofGroup.ToList());
				}

				var mat5FloorGroup = slabs.Where(s => s.Rsm && s.SlabType == SlabType.Floor)
						.GroupBy(s => s.ElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().Name},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        {"Percentage", (s.Sum(w => w.NetArea)/slabs.Where(sl => sl.SlabType == SlabType.Floor).Sum(w => w.NetArea)).ToString("P")},
                        {"Tier Level", s.First().RsmText}
                    });

				if (mat5FloorGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Floor", mat5FloorGroup.ToList());
				}

				var mat5BaselabGroup = slabs.Where(s => s.Rsm && s.SlabType == SlabType.BaseSlab)
						.GroupBy(s => s.ElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().Name},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        {"Percentage", (s.Sum(w => w.NetArea)/slabs.Where(sl => sl.SlabType == SlabType.BaseSlab).Sum(w => w.NetArea)).ToString("P")},
                        {"Tier Level", s.First().RsmText}
                    });

				if (mat5BaselabGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Foundation", mat5BaselabGroup.ToList());
				}

				var mat5ColumnGroup =
					columns.Where(c => c.Rsm).GroupBy(c => c.ElementNumber).Select(c => new Dictionary<string, object>
                {
                    {"Name", c.First().Name},
                    {"Element Number", c.Key},
                    {"Area", c.Sum(w => w.GrossSurfaceArea).ToString("N")},
                    {"Percentage", (c.Sum(w => w.GrossSurfaceArea)/columns.Sum(w => w.GrossSurfaceArea)).ToString("P")},
                    {"Tier Level", c.First().RsmText}
                });

				if (mat5ColumnGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Column", mat5ColumnGroup.ToList());
				}

				var mat5BeamGroup = beams.Where(c => c.Rsm).GroupBy(c => c.ElementNumber).Select(c => new Dictionary<string, object>
                {
                    {"Name", c.First().Name},
                    {"Element Number", c.Key},
                    {"Area", c.Sum(w => w.GrossSurfaceArea).ToString("N")},
                    {"Percentage", (c.Sum(w => w.GrossSurfaceArea)/beams.Sum(w => w.GrossSurfaceArea)).ToString("P")},
                    {"Tier Level", c.First().RsmText}
                });
				if (mat5BeamGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Beam", mat5BeamGroup.ToList());
				}

				var mat5StairGroup =
					stairs.Where(s => s.Rsm).GroupBy(s => s.ElementNumber).Select(s => new Dictionary<string, object>
                {
                    {"Name", s.First().Name},
                    {"Element Number", s.Key},
                    {"Area", s.Sum(sf => sf.Area).ToString("N")},
                    {"Percentage", (s.Sum(sf => sf.Area)/stairs.Sum(sf => sf.Area)).ToString("P")},
                    {"Tier Level", s.First().RsmText}
                });

				if (mat5StairGroup.Any())
				{
					mat5Items.Add("Responsibly Sourced Stair", mat5StairGroup.ToList());
				}
				if (!mat5Items.Any())
				{
					mat5Items.Add("Responsibly Sourced Materials", new List<Dictionary<string, object>>());
				}
				AddTable("Mat 5", mat5Items, ref pdfDoc);
			}

			if (documentCertficate.ContainsKey("Mat6a"))
			{
				var mat6AItems = new Dictionary<string, List<Dictionary<string, object>>>();

				var mat6AExternalWallGroup =
				externalWalls.Where(ew => ew.InsulationLayerThickness > 0)
					.GroupBy(ew => ew.InsulationMaterialElementNumber)
					.Select(ew => new Dictionary<string, object>
                    {
                        {"Name", ew.First().InsulationMaterialName},
                        {"Element Number", ew.Key},
                        {"Area", ew.Sum(w => w.NetSideArea).ToString("N")},
                        //{"Percentage", (ew.Sum(w => w.NetSideArea) / externalWalls.Sum(w => w.NetSideArea)).ToString("P")},
                        {"Green Guide Rating", ew.First().InsulationLayerMaterialPoint}
                    });

				mat6AItems.Add("External Wall Insulation", mat6AExternalWallGroup.ToList());

				var mat6ABaseSlabGroup = slabs.Where(s => s.InsulationLayerThickness > 0 && s.SlabType == SlabType.BaseSlab)
						.GroupBy(s => s.InsulationMaterialElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().InsulationMaterialName},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        //{"Percentage", (s.Sum(w => w.NetArea) / slabs.Where(sl => sl.SlabType == SlabType.BaseSlab).Sum(w => w.NetArea)).ToString("P")},
                        {"Green Guide Rating", s.First().InsulationLayerMaterialPoint}
                    });

				mat6AItems.Add("Ground Insulation", mat6ABaseSlabGroup.ToList());

				var mat6ARoofSlabGroup = slabs.Where(s => s.InsulationLayerThickness > 0 && s.SlabType == SlabType.Roof)
						.GroupBy(s => s.InsulationMaterialElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().InsulationMaterialName},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        //{"Percentage", (s.Sum(w => w.NetArea) / slabs.Where(sl => sl.SlabType == SlabType.Roof).Sum(w => w.NetArea)).ToString("P")},
                        {"Green Guide Rating", s.First().InsulationLayerMaterialPoint}
                    });

				mat6AItems.Add("Roof Insulation", mat6ARoofSlabGroup.ToList());
				if (!mat6AItems.Any())
				{
					mat6AItems.Add("Insulation", new List<Dictionary<string, object>>());
				}
				AddTable("Mat 6a", mat6AItems, ref pdfDoc);
			}

			if (documentCertficate.ContainsKey("Mat6b"))
			{
				var mat6BItems = new Dictionary<string, List<Dictionary<string, object>>>();
				var mat6BExternalWallGroup =
				externalWalls.Where(ew => ew.InsulationLayerThickness > 0 && ew.InsulationMaterialName.Contains("RSM"))
					.GroupBy(ew => ew.InsulationMaterialElementNumber)
					.Select(ew => new Dictionary<string, object>
                    {
                        {"Name", ew.Single().InsulationMaterialName},
                        {"Element Number", ew.Key},
                        {"Area", ew.Sum(w => w.NetSideArea).ToString("N")},
                        {"Percentage", (ew.Sum(w => w.NetSideArea) / externalWalls.Sum(w => w.NetSideArea)).ToString("P")},
                        {"Green Guide Rating", ew.Single().InsulationLayerMaterialPoint}
                    });
				if (mat6BExternalWallGroup.Any())
				{
					mat6BItems.Add("Responsibly Sourced External Wall Insulation", mat6BExternalWallGroup.ToList());
				}

				var mat6BBaseSlabGroup = slabs.Where(s => s.InsulationLayerThickness > 0 && s.SlabType == SlabType.BaseSlab && s.InsulationMaterialName.Contains("RSM"))
						.GroupBy(s => s.InsulationMaterialElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().InsulationMaterialName},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        {"Percentage", (s.Sum(w => w.NetArea) / slabs.Where(sl => sl.SlabType == SlabType.BaseSlab).Sum(w => w.NetArea)).ToString("P")},
                        {"Green Guide Rating", s.First().InsulationLayerMaterialPoint}
                    });
				if (mat6BBaseSlabGroup.Any())
				{
					mat6BItems.Add("Responsibly Sourced Ground Insulation", mat6BBaseSlabGroup.ToList());
				}

				var mat6BRoofSlabGroup = slabs.Where(s => s.InsulationLayerThickness > 0 && s.SlabType == SlabType.Roof && s.InsulationMaterialName.Contains("RSM"))
						.GroupBy(s => s.InsulationMaterialElementNumber)
						.Select(s => new Dictionary<string, object>
                    {
                        {"Name", s.First().InsulationMaterialName},
                        {"Element Number", s.Key},
                        {"Area", s.Sum(w => w.NetArea).ToString("N")},
                        {"Percentage", (s.Sum(w => w.NetArea) / slabs.Where(sl => sl.SlabType == SlabType.Roof).Sum(w => w.NetArea)).ToString("P")},
                        {"Green Guide Rating", s.First().InsulationLayerMaterialPoint}
                    });
				if (mat6BRoofSlabGroup.Any())
				{
					mat6BItems.Add("Responsibly Sourced Roof Insulation", mat6BRoofSlabGroup.ToList());
				}
				if (!mat6BItems.Any())
				{
					mat6BItems.Add("Responsibly Sourced Insulation", new List<Dictionary<string, object>>());
				}
				AddTable("Mat 6b", mat6BItems, ref pdfDoc);
			}

			if (documentCertficate.ContainsKey("Mat7"))
			{
				var mat7Items = new Dictionary<string, List<Dictionary<string, object>>>();
				mat7Items.Add("Designing for Robustness", new List<Dictionary<string, object>>());
				AddTable("Mat 7", mat7Items, ref pdfDoc);
			}

			pdfDoc.createPDF(pdfFilePath);
			pdfDoc = null;
		}

		private void AddTable(string tableName, Dictionary<string, List<Dictionary<string, object>>> items, ref pdfDocument doc)
		{
			var page = doc.addPage(predefinedPageSize.csA4Page);
			page.addText(tableName, 40, page.height - 30, doc.getFontReference(predefinedFont.csHelvetica), 8);

			var table = new pdfTable(doc, 1, pdfColor.Black, 2)
			{
				coordX = 40,
				coordY = page.height - 40
			};
			//tableName == "Mat 2a" || tableName == "Mat 7"
			if ( (items.Count == 1 && items.First().Value.Count < 1) || documentCertficate[tableName.Replace(" ", string.Empty)] == "No Credit")
			{
				table.tableHeader.addColumn(255);
				table.tableHeader.addColumn(255);
				if (items.Count > 0)
				{
					var nameRow = table.createRow();
					nameRow[0].addText(items.First().Key);
					table.addRow(nameRow);
				}
				var footerRow = table.createRow();
				footerRow[0].addText("Score");
				footerRow[1].addParagraph(documentCertficate[tableName.Replace(" ", string.Empty)], 10, predefinedAlignment.csLeft);
				table.addRow(footerRow);
				page.addTable(table);
				return;
			}

			var columnCount = items.First().Value != null && items.First().Value.Any() ? items.First().Value.First().Keys.Count + 1 : 2;

			for (var i = 0; i < columnCount; i++)
			{
				table.tableHeader.addColumn((int)Math.Floor((decimal)(510 / columnCount)));
			}

			foreach (var item in items)
			{
				if (item.Value.Count > 0)
				{
					var row = table.createRow();
					row[0].addParagraph(item.Key, 10, predefinedAlignment.csLeft);
					var i = 1;
					foreach (var key in item.Value[0].Keys)
					{
						row[i].addParagraph(key, 10, predefinedAlignment.csLeft);
						i++;
					}
					table.addRow(row);
					foreach (var elementGroup in item.Value)
					{
						row = table.createRow();
						row[0].addText((item.Value.IndexOf(elementGroup) + 1).ToString(CultureInfo.InvariantCulture));
						i = 1;
						foreach (var kvp in elementGroup)
						{
							if (kvp.Key == "Green Guide Rating")
							{
								var point = (decimal)kvp.Value;
								string grade;
								if (point > 2)
								{
									grade = "A+";
								}
								else if (point > 1)
								{
									grade = "A";
								}
								else if (point > (decimal)0.5)
								{
									grade = "B";
								}
								else if (point > (decimal)0.25)
								{
									grade = "C";
								}
								else if (point > 0)
								{
									grade = "D";
								}
								else
								{
									grade = "E";
								}
								row[i].addText(grade);
							}
							else if (kvp.Key == "Tier Level")
							{
								if (kvp.Value.ToString().Contains("Excellent"))
								{
									row[i].addText("Excellent");
								}
								else if (kvp.Value.ToString().Contains("Very Good"))
								{
									row[i].addText("Very Good");
								}
								else if (kvp.Value.ToString().Contains("Good"))
								{
									row[i].addText("Good");
								}
								else if (kvp.Value.ToString().Contains("Certified"))
								{
									row[i].addText("Certified EMS");
								}
								else if (kvp.Value.ToString().Contains("Verified"))
								{
									row[i].addText("Verified");
								}
								else
								{
									row[i].addText(string.Empty);
								}
							}
							else
							{
								row[i].addParagraph(kvp.Value != null ? kvp.Value.ToString() : string.Empty, 10, predefinedAlignment.csLeft);
							}
							i++;
						}
						table.addRow(row);
					}
				}
			}
			if (documentCertficate.ContainsKey(tableName.Replace(" ", string.Empty)))
			{
				var footerRow = table.createRow();
				footerRow[columnCount - 2].addText("Score");
				footerRow[columnCount - 1].addParagraph(documentCertficate[tableName.Replace(" ", string.Empty)], 10,
					predefinedAlignment.csLeft);
				table.addRow(footerRow);
			}
			page.addTable(table);
		}


		private void ReportClicked(object sender, EventArgs e)
		{
			var fileDialog = new SaveFileDialog { DefaultExt = ".pdf", OverwritePrompt = true, AddExtension = true };
			var fileDialogResult = fileDialog.ShowDialog();
			if (fileDialogResult == DialogResult.OK)
			{
				CreatePdf(fileDialog.FileName);
			}
		}
	}
}