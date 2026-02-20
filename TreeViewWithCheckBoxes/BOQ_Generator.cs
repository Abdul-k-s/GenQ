using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace GenQ
{
    /// <summary>
    /// Bill of Quantities Generator - calculates quantities for Revit elements
    /// Optimized for Revit 2025 API
    /// </summary>
    public class BOQ_Generator
    {
        #region Constants
        
        // Unit conversion constants (Revit internal units are feet-based)
        private const double FEET_TO_METERS = 0.3048;
        private const double SQFT_TO_SQM = 0.092903;
        private const double CUFT_TO_CUM = 0.0283168466;

        // Excluded categories for bounding box filtering (using HashSet for O(1) lookup)
        private static readonly HashSet<string> ExcludedCategories = new HashSet<string>
        {
            "Levels", "Grids", "Location Data", "RVT Links",
            "Doors", "Floors", "Roofs", "Generic Models", "Railings",
            "Structural Columns", "Structural Foundations", "Structural Framing",
            "Walls", "Windows"
        };

        #endregion

        #region Properties

        public bool AllGood { get; private set; } = true;
        public string ErrorMessage { get; private set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Processes all view types and calculates quantities based on unit type
        /// </summary>
        public List<FooViewModel> Sort(List<FooViewModel> ViewTypes, List<string> LevelName)
        {
            if (ViewTypes == null || LevelName == null || !LevelName.Any())
            {
                return ViewTypes ?? new List<FooViewModel>();
            }

            // Reset state
            AllGood = true;
            ErrorMessage = null;

            foreach (FooViewModel Type in ViewTypes)
            {
                if (!AllGood)
                    break;

                switch (Type.UnitType)
                {
                    case UnitTypeEnum.Length:
                        LengthCalc(Type, LevelName);
                        Type.Quantity = Type.Length;
                        break;
                    case UnitTypeEnum.Area:
                        AreaCalc(Type, LevelName);
                        Type.Quantity = Type.Area;
                        break;
                    case UnitTypeEnum.Volume:
                        VolumeCalc(Type, LevelName);
                        Type.Quantity = Type.Volume;
                        break;
                    case UnitTypeEnum.Perimeter:
                        PerimeterCalc(Type, LevelName);
                        Type.Quantity = Type.Perimeter;
                        break;
                    case UnitTypeEnum.Count:
                        CountCalc(Type, LevelName);
                        Type.Quantity = Type.Count;
                        break;
                }

                Type.Total = Type.Quantity * Type.Cost;
            }

            return ViewTypes;
        }

        #endregion

        #region Element Instance Retrieval

        /// <summary>
        /// Centralized helper to get element instances based on level selection
        /// </summary>
        private List<Element> GetElementInstances(Element type, List<string> levelName)
        {
            if (levelName == null || !levelName.Any())
                return new List<Element>();

            return levelName[0] != "All Levels" 
                ? GetInstanceListFromTypeByLevel(type, levelName) 
                : GetInstanceListFromType(type);
        }

        /// <summary>
        /// Gets all instances of a given element type from the document
        /// </summary>
        public List<Element> GetInstanceListFromType(Element type)
        {
            if (type?.Document == null)
                return new List<Element>();

            return new FilteredElementCollector(type.Document)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(instance => instance?.GetTypeId()?.Value == type.Id.Value)
                .ToList();
        }

        /// <summary>
        /// Gets all instances of a given element type filtered by levels
        /// </summary>
        public List<Element> GetInstanceListFromTypeByLevel(Element type, List<string> levelName)
        {
            if (type?.Document == null || levelName == null || !levelName.Any())
                return new List<Element>();

            Document doc = type.Document;
            List<Element> resultElements = new List<Element>();

            // Get and sort levels by elevation
            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();

            // Get level IDs for selected level names
            HashSet<ElementId> selectedLevelIds = new HashSet<ElementId>(
                levels.Where(l => levelName.Contains(l.Name)).Select(l => l.Id)
            );

            if (!selectedLevelIds.Any())
                return resultElements;

            // Find a valid 3D view for bounding box calculations
            View3D view3d = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => v != null && !v.IsTemplate && v.Title != "3D View: Analytical Model");

            if (view3d == null)
            {
                ErrorMessage = "No valid 3D view found for bounding box calculations.";
                MessageBox.Show(ErrorMessage);
                return resultElements;
            }

            // Process each selected level
            for (int i = 0; i < levels.Count; i++)
            {
                if (!selectedLevelIds.Contains(levels[i].Id))
                    continue;

                // Get elements by bounding box
                AddElementsByBoundingBox(type, levels, i, view3d, resultElements);

                // Get elements by level filter
                AddElementsByLevelFilter(type, levels[i], resultElements);
            }

            // Return deduplicated list
            return DeduplicateElements(resultElements);
        }

        /// <summary>
        /// Adds elements found within bounding box of the level
        /// </summary>
        private void AddElementsByBoundingBox(Element type, List<Level> levels, int levelIndex, View3D view3d, List<Element> results)
        {
            BoundingBoxXYZ box = levels[levelIndex].get_BoundingBox(view3d);
            if (box == null)
                return;

            XYZ dvd, dvdmax;

            if (levelIndex == levels.Count - 1)
            {
                // Top level - extend upward
                if (levelIndex == 0)
                    return; // Only one level, skip bounding box approach

                BoundingBoxXYZ box2 = levels[levelIndex - 1].get_BoundingBox(view3d);
                if (box2 == null)
                    return;

                XYZ diff = new XYZ(0, 0, (box.Max.Z - box.Min.Z) / 2.1);
                dvd = box.Min + diff;
                dvdmax = new XYZ(box.Max.X, box.Max.Y, box.Max.Z + (box.Max.Z - box2.Max.Z));
            }
            else
            {
                // Middle/bottom level
                BoundingBoxXYZ box2 = levels[levelIndex + 1].get_BoundingBox(view3d);
                if (box2 == null)
                    return;

                XYZ diff = new XYZ(0, 0, (box.Max.Z - box.Min.Z) / 2.1);
                XYZ diff2 = new XYZ(0, 0, (box2.Max.Z - box2.Min.Z) / 2.1);
                dvd = box.Min + diff;
                dvdmax = box2.Max - diff2;
            }

            Outline outline = new Outline(dvd, dvdmax);
            if (outline.IsEmpty)
                return;

            BoundingBoxIsInsideFilter filter = new BoundingBoxIsInsideFilter(outline);

            var elementsInBounds = new FilteredElementCollector(type.Document)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .Where(instance => instance?.GetTypeId()?.Value == type.Id.Value)
                .Where(e => e?.Category?.Name != null 
                    && !e.Category.Name.EndsWith(".dwg") 
                    && !ExcludedCategories.Contains(e.Category.Name))
                .ToList();

            results.AddRange(elementsInBounds);
        }

        /// <summary>
        /// Adds elements associated with a specific level using ElementLevelFilter
        /// </summary>
        private void AddElementsByLevelFilter(Element type, Level level, List<Element> results)
        {
            ElementLevelFilter levelFilter = new ElementLevelFilter(level.Id);

            var elementsOnLevel = new FilteredElementCollector(type.Document)
                .WherePasses(levelFilter)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(instance => instance?.GetTypeId()?.Value == type.Id.Value)
                .ToList();

            results.AddRange(elementsOnLevel);

            // Special handling for structural framing (beams)
            if (type.Category?.Name == "Structural Framing")
            {
                var beams = new FilteredElementCollector(type.Document)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .Where(e => e.LookupParameter("Reference Level")?.AsValueString() == level.Name)
                    .Where(instance => instance?.GetTypeId()?.Value == type.Id.Value)
                    .ToList();

                results.AddRange(beams);
            }
        }

        /// <summary>
        /// Removes duplicate elements using HashSet for O(1) lookup
        /// </summary>
        private List<Element> DeduplicateElements(List<Element> elements)
        {
            HashSet<long> seenIds = new HashSet<long>();
            List<Element> uniqueElements = new List<Element>();

            foreach (Element elem in elements)
            {
                if (elem != null && seenIds.Add(elem.Id.Value))
                {
                    uniqueElements.Add(elem);
                }
            }

            return uniqueElements;
        }

        #endregion

        #region Calculation Methods

        private FooViewModel PerimeterCalc(FooViewModel Type, List<string> LevelName)
        {
            try
            {
                List<Element> list = GetElementInstances(Type.Type, LevelName);
                double totalPerimeter = 0.0;

                foreach (Element elem in list)
                {
                    if (elem == null) continue;

                    Parameter parameter = elem.get_Parameter(BuiltInParameter.HOST_PERIMETER_COMPUTED);

                    if (parameter != null)
                    {
                        totalPerimeter += parameter.AsDouble();
                    }
                    else
                    {
                        SetError($"Cannot calculate {Type.Name} by Perimeter - parameter not found");
                        break;
                    }
                }

                if (AllGood)
                {
                    Type.Perimeter = Math.Round(totalPerimeter * FEET_TO_METERS, 2);
                }
            }
            catch (Exception ex)
            {
                SetError($"Error calculating perimeter for {Type.Name}: {ex.Message}");
            }

            return Type;
        }

        private FooViewModel VolumeCalc(FooViewModel Type, List<string> LevelName)
        {
            try
            {
                List<Element> list = GetElementInstances(Type.Type, LevelName);
                double totalVolume = 0.0;

                foreach (Element elem in list)
                {
                    if (elem == null) continue;

                    Parameter parameter = elem.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);

                    if (parameter != null)
                    {
                        totalVolume += parameter.AsDouble();
                    }
                    else
                    {
                        SetError($"Cannot calculate {Type.Name} by Volume - parameter not found");
                        break;
                    }
                }

                if (AllGood)
                {
                    Type.Volume = Math.Round(totalVolume * CUFT_TO_CUM, 2);
                }
            }
            catch (Exception ex)
            {
                SetError($"Error calculating volume for {Type.Name}: {ex.Message}");
            }

            return Type;
        }

        private FooViewModel AreaCalc(FooViewModel Type, List<string> LevelName)
        {
            try
            {
                List<Element> list = GetElementInstances(Type.Type, LevelName);
                double totalArea = 0.0;

                foreach (Element elem in list)
                {
                    if (elem == null) continue;

                    Parameter parameter = elem.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);

                    if (parameter != null)
                    {
                        totalArea += parameter.AsDouble();
                    }
                    else
                    {
                        SetError($"Cannot calculate {Type.Name} by Area - parameter not found");
                        break;
                    }
                }

                if (AllGood)
                {
                    Type.Area = Math.Round(totalArea * SQFT_TO_SQM, 2);
                }
            }
            catch (Exception ex)
            {
                SetError($"Error calculating area for {Type.Name}: {ex.Message}");
            }

            return Type;
        }

        private FooViewModel LengthCalc(FooViewModel Type, List<string> LevelName)
        {
            try
            {
                List<Element> list = GetElementInstances(Type.Type, LevelName);
                double totalLength = 0.0;

                foreach (Element elem in list)
                {
                    if (elem == null) continue;

                    Parameter parameter = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);

                    if (parameter != null)
                    {
                        totalLength += parameter.AsDouble();
                    }
                    else
                    {
                        SetError($"Cannot calculate {Type.Name} by Length - parameter not found");
                        break;
                    }
                }

                if (AllGood)
                {
                    Type.Length = Math.Round(totalLength * FEET_TO_METERS, 2);
                }
            }
            catch (Exception ex)
            {
                SetError($"Error calculating length for {Type.Name}: {ex.Message}");
            }

            return Type;
        }

        private FooViewModel CountCalc(FooViewModel Type, List<string> LevelName)
        {
            try
            {
                List<Element> list = GetElementInstances(Type.Type, LevelName);
                Type.Count = list.Count;
            }
            catch (Exception ex)
            {
                SetError($"Error calculating count for {Type.Name}: {ex.Message}");
            }

            return Type;
        }

        #endregion

        #region Helper Methods

        private void SetError(string message)
        {
            ErrorMessage = message;
            AllGood = false;
            MessageBox.Show(message);
        }

        #endregion
    }
}
