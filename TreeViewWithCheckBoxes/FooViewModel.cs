using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace GenQ
{
    public class FooViewModel : INotifyPropertyChanged
    {
        #region Data

        bool? _isChecked = false;
        FooViewModel _parent;
        public ElementType Type;

        #endregion // Data
        

        #region CreateFoos
        public static List<FooViewModel> Tree { get; set; }
        public List<FooViewModel> Getfromrevit(Document document, bool includelinks, List<String> LevelName)
        {
            List<Element> Types = new List<Element>();
            if (LevelName.Count == 0)
            {
                MessageBox.Show("Please Choose Levels to Show!");
            }
            else if (LevelName[0] != "All Levels")
            {
                Types = GetTypesByLevels(document, LevelName);
                if (includelinks)
                {
                    GetElementsfromRevitLinkByLevel(document, Types, LevelName);
                }
            }
            else
            {
                Types = GetTypes(document);

                if (includelinks)
                {
                    GetElementsfromRevitLink(document, Types);
                }

            }

            // Use HashSet for efficient distinct category collection
            HashSet<string> catNames = new HashSet<string>();
            foreach (ElementType Type in Types)
            {
                if (Type?.Category?.Name != null)
                {
                    catNames.Add(Type.Category.Name);
                }
            }

            List<string> typeNames = new List<string>();
            foreach (ElementType type in Types)
            {
                typeNames.Add(type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString());

            }

            FooViewModel root = new FooViewModel("All Categories") { IsInitiallySelected = true };
            
            // Pre-create category lookup dictionary for O(1) access
            Dictionary<string, FooViewModel> categoryLookup = new Dictionary<string, FooViewModel>();
            foreach (string cat in catNames.OrderBy(name => name))
            {
                FooViewModel categoryViewModel = new FooViewModel(cat);
                root.Children.Add(categoryViewModel);
                categoryLookup[cat] = categoryViewModel;
            }
            
            // Add types to their categories using dictionary lookup (O(1) instead of O(n))
            foreach (ElementType type in Types)
            {
                if (type?.Category?.Name == null) continue;
                
                Parameter nameParam = type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                if (nameParam == null) continue;
                
                string typeName = nameParam.AsString();
                string categoryName = type.Category.Name;
                
                if (categoryLookup.TryGetValue(categoryName, out FooViewModel categ))
                {
                    FooViewModel t = new FooViewModel(typeName);
                    t.Type = type;
                    categ.Children.Add(t);
                }
            }
            root.Initialize();
            Tree = new List<FooViewModel> { root };
            return Tree;
        }
        void GetElementsfromRevitLink(Document document, List<Element> Types)
        {
            List<Element> links = new FilteredElementCollector(document)
               .OfCategory(BuiltInCategory.OST_RvtLinks)
               .Where(fs => fs is RevitLinkInstance).ToList();

            List<Document> linkeddocs = new List<Document>();

            foreach (RevitLinkInstance link in links)
            {
                Document linkedoc = link.GetLinkDocument();
                linkeddocs.Add(linkedoc);

                //TaskDialog.Show("ss", $"{link.Name}   ,,  {link.Id}");//test
            }
            foreach (Document doc in linkeddocs)
            {
                if (doc == null)
                {
                    MessageBox.Show($"Link {links[linkeddocs.IndexOf(doc)].Name} is unloaded");
                }
                else
                {
                    List<Element> viewlinkedTypes = GetTypes(doc);
                    bool checked_ = true;
                    if (checked_)
                    {
                        Types.AddRange(viewlinkedTypes);
                    }
                }
            }
        }
        //List<Element> GetTypesByLevels(Document document, List<string> levelName)
        //{
        //    List<Element> ListOfTypesOnLevel = new List<Element>();

        //    FilteredElementCollector collector = new FilteredElementCollector(document);
        //    ICollection<Element> levels = collector.OfClass(typeof(Level)).ToElements();
        //    IEnumerable<Element> query = null;
        //    query = from element in collector where levelName.Contains(element.Name) select element;// Linq query

        //    List<ElementId> IDS = new List<ElementId>();
        //    foreach (Element level in query.ToList<Element>())
        //    {
        //        IDS.Add(level.Id);
        //    }

        //    FilteredElementCollector collector2 = new FilteredElementCollector(document);

        //    List<Element> TypesOnLevel = collector2.WhereElementIsNotElementType()
        //   .Where(e => e is Element && IDS.Contains(e.LevelId))
        //   .Select(e => e.GetTypeId())
        //   .Distinct()
        //   .Select(e => document.GetElement(e))
        //   .Where(e => e is ElementType && e.Category != null && e.Category.Name != "Levels" && e.Category.Name != "Grids" &&
        //   e.Category.Name != "Location Data" && e.Category.Name != "RVT Links")
        //   .ToList();

        //    // List<Element> beams = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFraming)
        //    //.WhereElementIsNotElementType().Where(e => levelName.Contains(e.LookupParameter("Reference Level").AsValueString())).Select(e => e.GetTypeId()).Distinct().Select(e => document.GetElement(e)).ToList();

        //    // TypesOnLevel.AddRange(beams);

        //    ListOfTypesOnLevel.AddRange(TypesOnLevel);


        //    return ListOfTypesOnLevel;
        //}
        List<Element> GetTypesByLevels(Document document, List<string> levelName)
        {
            List<Element> ListOfTypesOnLevel = new List<Element>();

            FilteredElementCollector collector = new FilteredElementCollector(document);
            List<Element> lvls = collector.OfClass(typeof(Level)).ToElements().ToList();
            List<Level> levels = new List<Level>();
            foreach (Element level in lvls)
            {
                levels.Add((Level)level);
            }

            levels = levels.OrderBy(p => p.Elevation).ToList();

            IEnumerable<Element> query = null;
            query = from element in collector where levelName.Contains(element.Name) select element;// Linq query

            List<ElementId> IDS = new List<ElementId>();
            foreach (Element level in query.ToList<Element>())
            {
                IDS.Add(level.Id);
            }

            FilteredElementCollector collector3d = new FilteredElementCollector(document).WhereElementIsNotElementType()
                    .OfClass(typeof(View3D));
            View3D view3d = null;
            foreach (View3D v in collector3d)
            {
                if (!v.IsTemplate && v.Title != "3D View: Analytical Model")
                {
                    view3d = v;
                    break;
                }
            }

            for (int i = 0; i < levels.Count; i++)
            {
                foreach (ElementId lvl in IDS)
                {
                    if (lvl == levels[i].Id)
                    {
                        XYZ dvd = new XYZ();
                        XYZ dvdmax = new XYZ();
                        if (i == levels.Count - 1)
                        {
                            BoundingBoxXYZ box = levels[i].get_BoundingBox(view3d);
                            BoundingBoxXYZ box2 = levels[i - 1].get_BoundingBox(view3d);
                            XYZ diff = new XYZ(0, 0, (box.Max.Z - box.Min.Z) / 2.1);
                            dvd = box.Min + diff;
                            XYZ dvdtemp = new XYZ(box.Max.X, box.Max.Y, box.Max.Z + (box.Max.Z - box2.Max.Z));
                            dvdmax = dvdtemp;
                        }
                        else
                        {
                            BoundingBoxXYZ box = levels[i].get_BoundingBox(view3d);
                            BoundingBoxXYZ box2 = levels[i + 1].get_BoundingBox(view3d);
                            XYZ diff = new XYZ(0, 0, (box.Max.Z - box.Min.Z) / 2.1);
                            XYZ diff2 = new XYZ(0, 0, (box2.Max.Z - box2.Min.Z) / 2.1);
                            dvd = box.Min + diff;
                            dvdmax = box2.Max - diff2;
                        }
                        // Create a BoundingBoxIsInside filter for Outline
                        Outline myOutLn = new Outline(dvd, dvdmax);
                        if (myOutLn.IsEmpty == true)
                        {
                            MessageBox.Show($"Level: {levels[i].Name} is Empty");
                        }
                        else
                        {
                            BoundingBoxIsInsideFilter filter = new BoundingBoxIsInsideFilter(myOutLn);

                            List<Element> TypesOnLevelBoun = new FilteredElementCollector(document).WhereElementIsNotElementType()
                           .WherePasses(filter)
                           .Select(e => e.GetTypeId())
                           .Distinct()
                           .Select(e => document.GetElement(e))
                           .Where(e => e is ElementType && e.Category != null && e.Category.Name.Substring(e.Category.Name.Length - 4) != ".dwg" && e.Category.Name != "Levels" && e.Category.Name != "Grids" &&
                           e.Category.Name != "Location Data" && e.Category.Name != "RVT Links" 
                           && e.Category.Name != "Doors" && e.Category.Name != "Floors" && e.Category.Name != "Roofs"
                           && e.Category.Name != "Generic Models" && e.Category.Name != "Railings" 
                           && e.Category.Name != "Structural Columns" && e.Category.Name != "Structural Foundations" 
                           && e.Category.Name != "Structural Framing" && e.Category.Name != "Walls" && e.Category.Name != "Windows")
                           .ToList();

                            ListOfTypesOnLevel.AddRange(TypesOnLevelBoun);
                        }

                    }
                }
            }
            FilteredElementCollector collector2 = new FilteredElementCollector(document);

            List<Element> TypesOnLevel = collector2.WhereElementIsNotElementType()
           .Where(e => e is Element && IDS.Contains(e.LevelId))
           .Select(e => e.GetTypeId())
           .Distinct()
           .Select(e => document.GetElement(e))
           .Where(e => e is ElementType && e.Category != null && e.Category.Name != "Levels" && e.Category.Name != "Grids" &&
           e.Category.Name != "Location Data" && e.Category.Name != "RVT Links")
           .ToList();

            List<Element> beams = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFraming)
           .WhereElementIsNotElementType().Where(e => levelName.Contains(e.LookupParameter("Reference Level").AsValueString())).Select(e => e.GetTypeId()).Distinct().Select(e => document.GetElement(e)).ToList();

            TypesOnLevel.AddRange(beams);

            ListOfTypesOnLevel.AddRange(TypesOnLevel);

            List<Element> UniqueTypes = ListOfTypesOnLevel.GroupBy(p => p.Name)
                          .Select(grp => grp.First())
                         .ToList();



            return UniqueTypes;
        }
        void GetElementsfromRevitLinkByLevel(Document document, List<Element> Types2, List<string> levelName)
        {
            List<Element> links = new FilteredElementCollector(document)
                         .OfCategory(BuiltInCategory.OST_RvtLinks)
                         .Where(fs => fs is RevitLinkInstance).ToList();

            Document linkedoc = null;
            foreach (RevitLinkInstance link in links)
            {
                linkedoc = link.GetLinkDocument();

                FilteredElementCollector collector = new FilteredElementCollector(linkedoc);
                List<Element> levels = collector.OfClass(typeof(Level)).Where(e => levelName.Contains(e.Name)).ToList();

                List<string> Names = new List<string>();
                foreach (Element level in levels)
                {
                    Names.Add(level.Name);
                }
                string Alllevels = $"{link.Name} doesn't contain level named: ";
                string diflevels = $"{link.Name} doesn't contain level named: ";
                foreach (string level in levelName)
                {
                    Alllevels = Alllevels + level + " ,";
                }
                foreach (string level in levelName.Except(Names))
                {
                    diflevels = diflevels + level + " ,";
                }

                if (levels.Count != 0)
                {
                    if (Names.Count == levelName.Count)
                    {
                        List<Element> viewlinkedTypes = GetTypesByLevels(linkedoc, levelName);
                        Types2.AddRange(viewlinkedTypes);
                    }
                    else
                    {
                        MessageBox.Show(diflevels);
                        List<Element> viewlinkedTypes = GetTypesByLevels(linkedoc, levelName);
                        Types2.AddRange(viewlinkedTypes);
                    }
                }
                else
                {
                    MessageBox.Show(Alllevels);
                }
            }
        }

        List<Element> GetTypes(Document document)
        {
            List<Element> Types = new FilteredElementCollector(document)
               .WhereElementIsNotElementType()
               .WhereElementIsViewIndependent()
               .Select(e => e.GetTypeId())
               .Distinct()
               .Select(e => document.GetElement(e))
               .Where(e => e is ElementType && e.Category != null && e.Category.Name != "Levels" && e.Category.Name != "Grids" &&
               e.Category.Name != "Location Data" && e.Category.Name != "RVT Links")
               .ToList();
            return Types;
        }

        FooViewModel(string name)
        {
            this.Name = name;
            this.Children = new List<FooViewModel>();
        }

        public FooViewModel()
        {
        }


        void Initialize()
        {
            foreach (FooViewModel child in this.Children)
            {
                child._parent = this;
                child.Initialize();
            }
        }

        #endregion // CreateFoos

        #region Properties

        public List<FooViewModel> Children { get; set; }

        public bool IsInitiallySelected { get; private set; }

        public string Name { get; private set; }

        #region IsChecked

        /// <summary>
        /// Gets/sets the state of the associated UI toggle (ex. CheckBox).
        /// The return value is calculated based on the check state of all
        /// child FooViewModels.  Setting this property to true or false
        /// will set all children to the same check state, and setting it 
        /// to any value will cause the parent to verify its check state.
        /// </summary>
        public bool? IsChecked
        {
            get { return _isChecked; }
            set { this.SetIsChecked(value, true, true); }
        }

        void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (value == _isChecked)
                return;

            _isChecked = value;

            if (updateChildren && _isChecked.HasValue)
                this.Children.ForEach(c => c.SetIsChecked(_isChecked, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyCheckState();

            this.OnPropertyChanged("IsChecked");
        }

        void VerifyCheckState()
        {
            bool? state = null;
            for (int i = 0; i < this.Children.Count; ++i)
            {
                bool? current = this.Children[i].IsChecked;
                if (i == 0)
                {
                    state = current;
                }
                else if (state != current)
                {
                    state = null;
                    break;
                }
            }
            this.SetIsChecked(state, false, true);
        }

        #endregion // IsChecked
        #region UnitType
        private UnitTypeEnum _UnitType;
        public UnitTypeEnum UnitType
        {
            get
            {

                return this._UnitType;
            }
            set
            {
                this.SetUnitType(value, true, true);
            }
        }
        void SetUnitType(UnitTypeEnum value, bool updateChildren, bool updateParent)
        {
            if (value == _UnitType)
                return;

            _UnitType = value;

            if (updateChildren)
                this.Children.ForEach(c => c.SetUnitType(_UnitType, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyUnitTypeState();

            this.OnPropertyChanged("UnitType");
        }
        void VerifyUnitTypeState()
        {
            UnitTypeEnum unit = UnitTypeEnum.None;
            List<UnitTypeEnum> units = new List<UnitTypeEnum>();
            for (int i = 0; i < this.Children.Count; ++i)
            {
                units.Add(this.Children[i].UnitType);

            }
            if (units.Distinct().Count() == 1)
            {
                unit = this.Children[0].UnitType;
            }
            this.SetUnitType(unit, false, true);
        }
        #endregion // UnitType
        #region Division
        private ReadMasterFormat _Division;
        public ReadMasterFormat Division
        {
            get { return this._Division; }
            set
            {
                this.SetDivision(value, true, true);
            }
        }
        void SetDivision(ReadMasterFormat value, bool updateChildren, bool updateParent)
        {
            if (value == _Division)
                return;

            _Division = value;

            if (updateChildren)
                this.Children.ForEach(c => c.SetDivision(_Division, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyDivisionState();

            this.OnPropertyChanged("Division");
        }
        void VerifyDivisionState()
        {
            ReadMasterFormat division = null;
            List<ReadMasterFormat> divisions = new List<ReadMasterFormat>();
            for (int i = 0; i < this.Children.Count; ++i)
            {
                divisions.Add(this.Children[i].Division);

            }
            if (divisions.Distinct().Count() == 1)
            {
                division = this.Children[0].Division;
            }
            this.SetDivision(division, false, true);
        }
        #endregion // Division
        #region Section
        private ReadMasterFormat _Section;
        public ReadMasterFormat Section
        {
            get { return this._Section; }
            set
            {
                this.SetSection(value, true, true);
            }
        }
        void SetSection(ReadMasterFormat value, bool updateChildren, bool updateParent)
        {
            if (value == _Section)
                return;

            _Section = value;

            if (updateChildren && this.Section != null)
                this.Children.ForEach(c => c.SetSection(_Section, true, false));

            if (updateParent && _parent != null)
                _parent.VerifySectionState();

            this.OnPropertyChanged("Section");
        }
        void VerifySectionState()
        {
            ReadMasterFormat section = null;
            List<ReadMasterFormat> sections = new List<ReadMasterFormat>();
            for (int i = 0; i < this.Children.Count; ++i)
            {
                sections.Add(this.Children[i].Section);

            }
            if (sections.Distinct().Count() == 1)
            {
                section = this.Children[0].Division;
            }
            this.SetSection(section, false, true);
        }
        #endregion // Section
        #region Cost
        private Double _Cost;
        public Double Cost
        {
            get { return this._Cost; }
            set
            {
                this.SetCost(value, true, true);
            }
        }
        void SetCost(Double value, bool updateChildren, bool updateParent)
        {
            if (value == _Cost)
                return;

            _Cost = value;

            if (updateChildren)
                this.Children.ForEach(c => c.SetCost(_Cost, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyCostState();

            this.OnPropertyChanged("Cost");
        }
        void VerifyCostState()
        {
            Double cost=0;
            List<Double> costs = new List<Double>();
            for (int i = 0; i < this.Children.Count; ++i)
            {
                costs.Add(this.Children[i].Cost);

            }
            if (costs.Distinct().Count() == 1)
            {
                cost = this.Children[0].Cost;
            }
            this.SetCost(cost, false, true);
        }
        #endregion // Cost

        #region Description
        private String _Description;
        public String Description
        {
            get { return this._Description; }
            set
            {
                
                if (value == null) { this.SetDescription(value, true, true); }
                else{ this.SetDescription(value.TrimStart(), true, true); }

            }
        }
        void SetDescription(String value, bool updateChildren, bool updateParent)
        {
            if (value == _Description)
                return;

            _Description = value;

            if (updateChildren)
                this.Children.ForEach(c => c.SetDescription(_Description, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyDescriptionState();

            this.OnPropertyChanged("Description");
        }
        void VerifyDescriptionState()
        {
            String description = "Add Description";
            List<String> descriptions = new List<String>();
            for (int i = 0; i < this.Children.Count; ++i)
            {
                descriptions.Add(this.Children[i].Description);

            }
            if (descriptions.Distinct().Count() == 1)
            {
                description = this.Children[0].Description;
            }
            this.SetDescription(description.TrimStart(), false, true);
        }
        #endregion // Description
        #region Area
        private Double _Area; // field
        public Double Area   // property
        {
            get { return _Area; }   // get method
            set { _Area = value; }  // set method
        }
        #endregion
        #region Volume
        private Double _Volume; // field
        public Double Volume   // property
        {
            get { return _Volume; }   // get method
            set { _Volume = value; }  // set method
        }
        #endregion
        #region Perimeter
        private Double _Perimeter; // field
        public Double Perimeter   // property
        {
            get { return _Perimeter; }   // get method
            set { _Perimeter = value; }  // set method
        }
        #endregion
        #region Length
        private Double _Length; // field
        public Double Length   // property
        {
            get { return _Length; }   // get method
            set { _Length = value; }  // set method
        }
        #endregion
        #region Count
        private Double _Count; // field
        public Double Count   // property
        {
            get { return _Count; }   // get method
            set { _Count = value; }  // set method
        }
        #endregion
        #region Quantity
        private Double _Quantity; // field
        public Double Quantity   // property
        {
            get { return _Quantity; }   // get method
            set { _Quantity = value; }  // set method
        }
        #endregion
        #region Total
        private Double _Total; // field
        public Double Total   // property
        {
            get { return _Total; }   // get method
            set { _Total = value; }  // set method
        }
        #endregion
        #endregion

        #region INotifyPropertyChanged Members

        void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
}
