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

using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Ecf.Excel.Xunit
{
    /// <summary>
    /// Integration tests for <see cref="XlsExportManager"/> and <see cref="CsvExportManager"/>.
    /// </summary>
    public class IntegrationTest
    {
        [Fact]
        public async Task TestXlsExport()
        {
            var ecfFolder = new DirectoryInfo(Path.Combine(GetOutputFolder(), "XlsExport"));
            if (!ecfFolder.Exists)
            {
                ecfFolder.Create();
            };

            var xlsFile = Path.Combine(GetOutputFolder(), "Assets", "test.xlsx");
            var cfgFile = Path.Combine(GetOutputFolder(), "Assets", "test.config.json");

            var csvConfig = await ConfigurationManager.LoadFromFile(cfgFile);

            csvConfig.EcfExport.TargetFolderName = ecfFolder.FullName;
            csvConfig.EcfExport.SourceProvider = EcfSourceProvider.Xlsx;
            csvConfig.EcfExport.SourceFileName = xlsFile;

            var exportManager = new XlsExportManager(csvConfig);

            await exportManager.Execute();
        }

        [Fact]
        public async Task TestCsvExport()
        {
            var ecfFolder = new DirectoryInfo(Path.Combine(GetOutputFolder(), "CsvExport"));
            if (!ecfFolder.Exists)
            {
                ecfFolder.Create();
            };

            var csvFile = Path.Combine(GetOutputFolder(), "Assets", "test.csv");
            var cfgFile = Path.Combine(GetOutputFolder(), "Assets", "test.config.json");

            var csvConfig = await ConfigurationManager.LoadFromFile(cfgFile);

            csvConfig.EcfExport.TargetFolderName = ecfFolder.FullName;
            csvConfig.EcfExport.SourceProvider = EcfSourceProvider.Csv;
            csvConfig.EcfExport.SourceFileName = csvFile;

            var exportManager = new CsvExportManager(csvConfig);

            await exportManager.Execute();
        }

        private string GetOutputFolder()
        {
            // Get the full location of the assembly
            string assemblyPath = System.Reflection.Assembly.GetAssembly(typeof(IntegrationTest)).Location;

            // Get the folder that's in
            return Path.GetDirectoryName(assemblyPath);
        }
    }
}
