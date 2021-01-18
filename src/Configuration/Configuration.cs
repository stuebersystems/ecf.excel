#region ENBREA - Copyright (C) 2021 STÜBER SYSTEMS GmbH
/*    
 *    ENBREA
 *    
 *    Copyright (C) 2021 STÜBER SYSTEMS GmbH
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

using System.Linq;

namespace Ecf.Excel
{
    public class Configuration
    {
        public EcfExportOptions EcfExport { get; set; } = new EcfExportOptions();

        public string GetCsvHeaderName(string ecfHeaderName)
        {
            var mapping = EcfExport?.CsvMappings.FirstOrDefault(x => x.ToHeader == ecfHeaderName);
            if (mapping != null)
            {
                return mapping.FromHeader;
            }
            else
            {
                return ecfHeaderName;
            }
        }

        public string GetXlsColumnName(string ecfHeaderName)
        {
            var mapping = EcfExport?.XlsMappings.FirstOrDefault(x => x.ToHeader == ecfHeaderName);
            if (mapping != null)
            {
                return mapping.FromHeader;
            }
            else
            {
                return null;
            }
        }
    }
}
