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

namespace Ecf.Excel
{
    public class ExportSubject
    {
        public readonly string Code;
        public readonly string Id;

        public ExportSubject(Configuration config, CsvTableReader csvTableReader, string headerName)
        {
            if (csvTableReader.TryGetValue(config.GetCsvHeaderName(headerName), out Id))
            {
                Code = Id;
            }
        }

        public ExportSubject(Configuration config, XlsReader xlsReader, string headerName)
        {
            if (xlsReader.TryGetValue(config.GetXlsColumnName(headerName), out Id))
            {
                Code = Id;
            }
        }
    }
}
