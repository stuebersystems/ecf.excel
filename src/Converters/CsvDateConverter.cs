#region ENBREA - Copyright (C) 2020 STÜBER SYSTEMS GmbH
/*    
 *    ENBREA
 *    
 *    Copyright (C) 2020 STÜBER SYSTEMS GmbH
 *
 *    This program is free software: you can redistribute it and/or modify
 *    it under the terms of the GNU Affero General Public License, version 3,
 *    as published by the Free Software Foundation.
 *
 *    This program is distributed in the hope that it will be useful,
 *    but WITHOUT ANY WARRANTY; without even the implied warranty of
 *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 *    GNU Affero General Public License for more details.
 *
 *    You should have received a copy of the GNU Affero General Public License
 *    along with this program. If not, see <http://www.gnu.org/licenses/>.
 *
 */
#endregion

using Enbrea.Csv;
using System;
using System.Globalization;

namespace Ecf.Excel
{
    /// <summary>
    /// Implementation of a DateTime converter to or from CSV
    /// </summary>
    public class CsvDateConverter : CsvDefaultConverter
    {
        private readonly string[] _formats =
        {
            "yyyy-MM-dd",
            "dd.MM.yyyy"
        };

        public CsvDateConverter() : 
            base(typeof(DateTime), CultureInfo.InvariantCulture)
        {
        }

        public override object FromString(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            else
            {
                return DateTime.ParseExact(value, _formats, CultureInfo, DateTimeStyles.None);
            }
        }

        public override string ToString(object value)
        {
            if ((value != null) && (value is DateTime dateTimeValue))
            {
                return dateTimeValue.ToString(_formats[0], CultureInfo);
            }
            else
            {
                return base.ToString(value);
            }
        }
    }
}