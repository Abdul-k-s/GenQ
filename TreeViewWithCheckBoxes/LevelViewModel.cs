using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace GenQ
{
    public class LevelViewModel : INotifyPropertyChanged
    {
        #region Data

        bool? _isChecked = false;
        LevelViewModel _parent;

        #endregion // Data
        #region CreateFoos
        public static List<LevelViewModel> lvls { get; set; }
        public List<LevelViewModel> Getfromrevit(Document document)
        {
            List<string> Levels = new List<string>();
            Levels = GetLevelsNames(document);

            LevelViewModel root = new LevelViewModel("All Levels") { IsInitiallySelected = true , IsChecked=true };
            foreach (String Level in Levels.Distinct().OrderBy(name => name).ToList())
            {
                root.Children.Add(new LevelViewModel(Level));
            }
            root.Initialize();
            lvls = new List<LevelViewModel> {root};
            return lvls;
        }
        public List<String> GetLevelsNames(Document doc)
        {
            List<string> Lvls = new List<string>();
            List<Element> levels = new FilteredElementCollector(doc)
                                              .OfCategory(BuiltInCategory.OST_Levels)
                                              .WhereElementIsNotElementType()
                                              .ToElements().ToList();
            foreach (Element level in levels)
            {
                Lvls.Add(level.Name);
            }
            return Lvls;
        }

        LevelViewModel(string name)
        {
            this.Name = name;
            this.Children = new List<LevelViewModel>();
        }

        public LevelViewModel()
        {
        }


        void Initialize()
        {
            foreach (LevelViewModel child in this.Children)
            {
                child._parent = this;
                child.Initialize();
            }
        }

        #endregion // CreateFoos

        #region Properties

        public List<LevelViewModel> Children { get; set; }

        public bool IsInitiallySelected { get; private set; }

        public string Name { get; private set; }

        #region IsChecked
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
