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

using ClosedXML.Excel;
using Enbrea.Csv;
using Enbrea.Ecf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ecf.Excel
{
    public class XlsExportManager : CustomManager
    {
        private int _recordCounter = 0;
        private int _tableCounter = 0;

        public XlsExportManager(
            Configuration config,
            CancellationToken cancellationToken = default,
            EventWaitHandle cancellationEvent = default)
            : base(config, cancellationToken, cancellationEvent)
        {
        }

        public async override Task Execute()
        {
            using var xlsStream = new FileStream(_config.EcfExport.SourceFileName, FileMode.Open, FileAccess.Read, FileShare.None);
            using var xlsDocument = new XLWorkbook(xlsStream, XLEventTracking.Disabled);
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
                await Execute(xlsDocument, EcfTables.Subjects, async (r, w, h) => await ExportSubjects(r, w, h));
                await Execute(xlsDocument, EcfTables.SchoolClasses, async (r, w, h) => await ExportSchoolClasses(r, w, h));
                await Execute(xlsDocument, EcfTables.Students, async (r, w, h) => await ExportStudents(r, w, h));
                await Execute(xlsDocument, EcfTables.StudentSchoolClassAttendances, async (r, w, h) => await ExportStudentSchoolClassAttendances(r, w, h));
                await Execute(xlsDocument, EcfTables.StudentSubjects, async (r, w, h) => await ExportStudentSubjects(r, w, h));

                // Report status
                Console.WriteLine($"[Extracting] {_tableCounter} table(s) and {_recordCounter} record(s) extracted");
            }
            catch
            {
                // Report error 
                Console.WriteLine();
                Console.WriteLine($"[Error] Extracting failed. Only {_tableCounter} table(s) and {_recordCounter} record(s) extracted");
                throw;
            }
        }

        private async Task Execute(IXLWorkbook xlsDocument, string ecfTableName, Func<XlsReader, EcfTableWriter, string[], Task<int>> action)
        {
            if (ShouldExportTable(ecfTableName, out var ecfFile))
            {
                // Report status
                Console.WriteLine($"[Extracting] [{ecfTableName}] Start...");

                // Create Excel Reader
                var xlsReader = new XlsReader(xlsDocument, _config?.EcfExport?.XlsSheetName, _config?.EcfExport?.XlsFirstRowNumber, _config?.EcfExport?.XlsLastRowNumber);

                // Generate ECF file name
                var ecfFileName = Path.ChangeExtension(Path.Combine(_config.EcfExport.TargetFolderName, ecfTableName), "csv");

                // Create ECF file for export
                using var ecfWriterStream = new FileStream(ecfFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None);

                // Create ECF Writer
                using var ecfWriter = new CsvWriter(ecfWriterStream, Encoding.UTF8);

                // Call table specific action
                var ecfRecordCounter = await action(xlsReader, new EcfTableWriter(ecfWriter), ecfFile?.Headers);

                // Inc counters
                _recordCounter += ecfRecordCounter;
                _tableCounter++;

                // Report status
                Console.WriteLine($"[Extracting] [{ecfTableName}] {ecfRecordCounter} record(s) extracted");
            }
        }

        private async Task<int> ExportSchoolClasses(XlsReader xlsReader, EcfTableWriter ecfTableWriter, string[] ecfHeaders)
        {
            var ecfCache = new HashSet<string>();
            var ecfRecordCounter = 0;

            if (ecfHeaders != null && ecfHeaders.Length > 0)
            {
                await ecfTableWriter.WriteHeadersAsync(ecfHeaders);
            }
            else
            {
                await ecfTableWriter.WriteHeadersAsync(
                    EcfHeaders.Id,
                    EcfHeaders.Code);
            }

            while (xlsReader.ReadLine())
            {
                var schoolClass = new ExportSchoolClass(_config, xlsReader);

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

        private async Task<int> ExportStudents(XlsReader xlsReader, EcfTableWriter ecfTableWriter, string[] ecfHeaders)
        {
            var ecfCache = new HashSet<string>();
            var ecfRecordCounter = 0;

            if (ecfHeaders != null && ecfHeaders.Length > 0)
            {
                await ecfTableWriter.WriteHeadersAsync(ecfHeaders);
            }
            else
            {
                await ecfTableWriter.WriteHeadersAsync(
                    EcfHeaders.Id,
                    EcfHeaders.LastName,
                    EcfHeaders.FirstName,
                    EcfHeaders.MiddleName,
                    EcfHeaders.NickName,
                    EcfHeaders.Salutation,
                    EcfHeaders.Gender,
                    EcfHeaders.Birthdate);
            }

            while (xlsReader.ReadLine())
            {
                var student = new ExportStudent(_config, xlsReader);

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

        private async Task<int> ExportStudentSchoolClassAttendances(XlsReader xlsReader, EcfTableWriter ecfTableWriter, string[] ecfHeaders)
        {
            var ecfRecordCounter = 0;

            if (ecfHeaders != null && ecfHeaders.Length > 0)
            {
                await ecfTableWriter.WriteHeadersAsync(ecfHeaders);
            }
            else
            {
                await ecfTableWriter.WriteHeadersAsync(
                    EcfHeaders.StudentId,
                    EcfHeaders.SchoolClassId);
            }

            while (xlsReader.ReadLine())
            {
                var student = new ExportStudent(_config, xlsReader);
                var schoolClass = new ExportSchoolClass(_config, xlsReader);

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

        private async Task<int> ExportStudentSubjects(XlsReader xlsReader, EcfTableWriter ecfTableWriter, string[] ecfHeaders)
        {
            var ecfRecordCounter = 0;

            if (ecfHeaders != null && ecfHeaders.Length > 0)
            {
                await ecfTableWriter.WriteHeadersAsync(ecfHeaders);
            }
            else
            {
                await ecfTableWriter.WriteHeadersAsync(
                    EcfHeaders.StudentId,
                    EcfHeaders.SchoolClassId,
                    EcfHeaders.SubjectId);
            }

            while (xlsReader.ReadLine())
            {
                var student = new ExportStudent(_config, xlsReader);
                var schoolClass = new ExportSchoolClass(_config, xlsReader);

                if (!string.IsNullOrEmpty(schoolClass.Id))
                {
                    for (int i = 1; i < 20; i++)
                    {
                        var subject = new ExportSubject(_config, xlsReader, $"Fach{i}");

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

        private async Task<int> ExportSubjects(XlsReader xlsReader, EcfTableWriter ecfTableWriter, string[] ecfHeaders)
        {
            var ecfCache = new HashSet<string>();
            var ecfRecordCounter = 0;

            if (ecfHeaders != null && ecfHeaders.Length > 0)
            {
                await ecfTableWriter.WriteHeadersAsync(ecfHeaders);
            }
            else
            {
                await ecfTableWriter.WriteHeadersAsync(
                    EcfHeaders.Id,
                    EcfHeaders.Code);
            }

            while (xlsReader.ReadLine())
            {
                for (int i = 1; i < 20; i++)
                {
                    var subject = new ExportSubject(_config, xlsReader, $"Fach{i}");

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
    }
}
