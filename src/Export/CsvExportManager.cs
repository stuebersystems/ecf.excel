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
using Enbrea.Ecf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ecf.Excel
{
    public class CsvExportManager : CustomManager
    {
        private int _recordCounter = 0;
        private int _tableCounter = 0;

        public CsvExportManager(
            Configuration config,
            CancellationToken cancellationToken = default,
            EventWaitHandle cancellationEvent = default)
            : base(config, cancellationToken, cancellationEvent)
        {
        }

        public async override Task Execute(bool ThrowExecptions = false)
        {
            try
            {
                // Init counters
                _tableCounter = 0;
                _recordCounter = 0;

                // Report status
                Console.WriteLine();
                Console.WriteLine("[Extracting] Start...");

                // Preperation
                PrepareExportFolder();

                // Education
                await Execute(EcfTables.Subjects, async (r, w) => await ExportSubjects(r, w));
                await Execute(EcfTables.SchoolClasses, async (r, w) => await ExportSchoolClasses(r, w));
                await Execute(EcfTables.Students, async (r, w) => await ExportStudents(r, w));
                await Execute(EcfTables.StudentSchoolClassAttendances, async (r, w) => await ExportStudentSchoolClassAttendances(r, w));
                await Execute(EcfTables.StudentSubjects, async (r, w) => await ExportStudentSubjects(r, w));

                // Report status
                Console.WriteLine($"[Extracting] {_tableCounter} table(s) and {_recordCounter} record(s) extracted");
            }
            catch (Exception ex)
            {
                if (!ThrowExecptions)
                {
                    // Report error 
                    Console.WriteLine();
                    Console.WriteLine($"[Error] Extracting failed. Only {_tableCounter} table(s) and {_recordCounter} record(s) extracted");
                    Console.WriteLine($"[Error] Reason: {ex.Message}");
                }
                else
                {
                    throw;
                }
            }
        }

        private async Task Execute(string ecfTableName, Func<CsvTableReader, EcfTableWriter, Task<int>> action)
        {
            EcfExportFile ecfFile = _config.EcfExport?.Files?.FirstOrDefault(x => x.Name.ToLower() == ecfTableName.ToLower());
            if (ecfFile != null)
            {
                // Report status
                Console.WriteLine($"[Extracting] [{ecfTableName}] Start...");

                // Open CSV file for import
                using var csvReaderStream = new FileStream(_config.EcfExport.SourceFileName, FileMode.Open, FileAccess.Read, FileShare.None);

                // Create CSV Reader
                using var csvReader = new CsvReader(csvReaderStream, Encoding.UTF8);

                // Create CSV Table Reader
                var csvTableReader = new CsvTableReader(csvReader, new CsvConverterResolver());

                // Generate ECF file name
                var ecfFileName = Path.ChangeExtension(Path.Combine(_config.EcfExport.FolderName, ecfTableName), "csv");

                // Create ECF file for export
                using var ecfWriterStream = new FileStream(ecfFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None);

                // Create ECF Writer
                using var ecfWriter = new CsvWriter(ecfWriterStream, Encoding.UTF8);

                // Call table specific action
                var ecfRecordCounter = await action(csvTableReader, new EcfTableWriter(ecfWriter));

                // Inc counters
                _recordCounter += ecfRecordCounter;
                _tableCounter++;

                // Report status
                Console.WriteLine($"[Extracting] [{ecfTableName}] {ecfRecordCounter} record(s) extracted");
            }
        }

        private async Task<int> ExportSchoolClasses(CsvTableReader csvTableReader, EcfTableWriter ecfTableWriter)
        {
            var ecfCache = new HashSet<string>();
            var ecfRecordCounter = 0;

            await csvTableReader.ReadHeadersAsync();

            await ecfTableWriter.WriteHeadersAsync(
                EcfHeaders.Id,
                EcfHeaders.Code);

            while (await csvTableReader.ReadAsync() > 0)
            {
                var schoolClass = new ExportSchoolClass(_config, csvTableReader);

                if (!string.IsNullOrEmpty(schoolClass.Id) && !ecfCache.Contains(schoolClass.Id))
                {
                    ecfTableWriter.SetValue(EcfHeaders.Id, schoolClass.Id);
                    ecfTableWriter.SetValue(EcfHeaders.Code, schoolClass.Code);

                    await ecfTableWriter.WriteAsync();

                    ecfCache.Add(schoolClass.Id);
                    ecfRecordCounter++;
                }
            }

            return ecfRecordCounter;
        }

        private async Task<int> ExportStudents(CsvTableReader csvTableReader, EcfTableWriter ecfTableWriter)
        {
            var ecfCache = new HashSet<string>();
            var ecfRecordCounter = 0;

            await csvTableReader.ReadHeadersAsync();

            await ecfTableWriter.WriteHeadersAsync(
                EcfHeaders.Id,
                EcfHeaders.LastName,
                EcfHeaders.FirstName,
                EcfHeaders.MiddleName,
                EcfHeaders.NickName,
                EcfHeaders.Salutation,
                EcfHeaders.Gender,
                EcfHeaders.Birthdate);

            while (await csvTableReader.ReadAsync() > 0)
            {
                var student = new ExportStudent(_config, csvTableReader);

                if (!ecfCache.Contains(student.Id))
                {
                    ecfTableWriter.SetValue(EcfHeaders.Id, student.Id);
                    ecfTableWriter.SetValue(EcfHeaders.LastName, student.LastName);
                    ecfTableWriter.SetValue(EcfHeaders.FirstName, student.FirstName);
                    ecfTableWriter.SetValue(EcfHeaders.MiddleName, student.MiddleName);
                    ecfTableWriter.SetValue(EcfHeaders.NickName, student.NickName);
                    ecfTableWriter.SetValue(EcfHeaders.Salutation, student.Salutation);
                    ecfTableWriter.SetValue(EcfHeaders.Gender, student.Gender);
                    ecfTableWriter.SetValue(EcfHeaders.Birthdate, student.BirthDate);

                    await ecfTableWriter.WriteAsync();

                    ecfCache.Add(student.Id);
                    ecfRecordCounter++;
                }
            }

            return ecfRecordCounter;
        }

        private async Task<int> ExportStudentSchoolClassAttendances(CsvTableReader csvTableReader, EcfTableWriter ecfTableWriter)
        {
            var ecfRecordCounter = 0;

            await csvTableReader.ReadHeadersAsync();

            await ecfTableWriter.WriteHeadersAsync(
                EcfHeaders.StudentId,
                EcfHeaders.SchoolClassId);

            while (await csvTableReader.ReadAsync() > 0)
            {
                var student = new ExportStudent(_config, csvTableReader);
                var schoolClass = new ExportSchoolClass(_config, csvTableReader);

                if (!string.IsNullOrEmpty(schoolClass.Id))
                {
                    ecfTableWriter.SetValue(EcfHeaders.StudentId, student.Id);
                    ecfTableWriter.SetValue(EcfHeaders.SchoolClassId, schoolClass.Id);

                    await ecfTableWriter.WriteAsync();

                    ecfRecordCounter++;
                }
            }

            return ecfRecordCounter;
        }

        private async Task<int> ExportStudentSubjects(CsvTableReader csvTableReader, EcfTableWriter ecfTableWriter)
        {
            var ecfRecordCounter = 0;

            await csvTableReader.ReadHeadersAsync();

            await ecfTableWriter.WriteHeadersAsync(
                EcfHeaders.StudentId,
                EcfHeaders.SchoolClassId,
                EcfHeaders.SubjectId);

            while (await csvTableReader.ReadAsync() > 0)
            {
                var student = new ExportStudent(_config, csvTableReader);
                var schoolClass = new ExportSchoolClass(_config, csvTableReader);

                if (!string.IsNullOrEmpty(schoolClass.Id))
                {
                    for (int i = 1; i < 20; i++)
                    {
                        var subject = new ExportSubject(_config, csvTableReader, $"Fach{i}");

                        if (!string.IsNullOrEmpty(subject.Id))
                        {
                            ecfTableWriter.SetValue(EcfHeaders.StudentId, student.Id);
                            ecfTableWriter.SetValue(EcfHeaders.SchoolClassId, schoolClass.Id);
                            ecfTableWriter.SetValue(EcfHeaders.SubjectId, subject.Id);

                            await ecfTableWriter.WriteAsync();

                            ecfRecordCounter++;
                        }
                    }
                }
            }

            return ecfRecordCounter;
        }

        private async Task<int> ExportSubjects(CsvTableReader csvTableReader, EcfTableWriter ecfTableWriter)
        {
            var ecfCache = new HashSet<string>();
            var ecfRecordCounter = 0;

            await csvTableReader.ReadHeadersAsync();

            await ecfTableWriter.WriteHeadersAsync(
                EcfHeaders.Id,
                EcfHeaders.Code);

            while (await csvTableReader.ReadAsync() > 0)
            {
                for (int i = 1; i < 20; i++)
                {
                    var subject = new ExportSubject(_config, csvTableReader, $"Fach{i}");

                    if (!string.IsNullOrEmpty(subject.Id) && !ecfCache.Contains(subject.Id))
                    {
                        ecfTableWriter.SetValue(EcfHeaders.Id, subject.Id);
                        ecfTableWriter.SetValue(EcfHeaders.Code, subject.Code);

                        await ecfTableWriter.WriteAsync();

                        ecfCache.Add(subject.Id);
                        ecfRecordCounter++;
                    }
                }
            }

            return ecfRecordCounter;
        }

        private void PrepareExportFolder()
        {
            if (Directory.Exists(_config.EcfExport.FolderName))
            {
                foreach (var fileName in Directory.EnumerateFiles(_config.EcfExport.FolderName, "*.csv"))
                {
                    File.Delete(fileName);
                }
            }
            else
            {
                Directory.CreateDirectory(_config.EcfExport?.FolderName);
            }
        }
    }
}
