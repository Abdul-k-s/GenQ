using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Data;
using System.Windows.Markup;

namespace GenQ
{
    public class UnitEnumConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var convert = new Dictionary<UnitTypeEnum, string>()
        {
            {UnitTypeEnum.None, "None"},
            {UnitTypeEnum.Area, "Area"},
            {UnitTypeEnum.Length, "Length"},
            {UnitTypeEnum.Volume, "Volume"},
            {UnitTypeEnum.Perimeter, "Perimeter"},
            {UnitTypeEnum.Count, "Count"}
        };

            return convert;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
