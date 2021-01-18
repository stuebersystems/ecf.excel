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

using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace Ecf.Excel
{
    public class EcfExportOptions : EcfOptions
    {
        public ICollection<CsvMapping> CsvMappings { get; set; } = new List<CsvMapping>();
        public ICollection<EcfExportFile> Files { get; set; } = new List<EcfExportFile>();
        public string SourceFileName { get; set; }
        [JsonConverter(typeof(JsonStringEnumConverter))]
        public EcfSourceProvider SourceProvider { get; set; } = EcfSourceProvider.Csv;
        public string TargetFolderName { get; set; }
        public int? XlsFirstRowNumber { get; set; }
        public int? XlsLastRowNumber { get; set; }
        public ICollection<XlsMapping> XlsMappings { get; set; } = new List<XlsMapping>();
        public string XlsSheetName { get; set; }
    }
}
