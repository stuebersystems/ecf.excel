﻿#region ENBREA - Copyright (C) 2021 STÜBER SYSTEMS GmbH
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

using Enbrea.Csv;
using Enbrea.Ecf;
using Enbrea.GuidFactory;
using System;

namespace Ecf.Excel
{
    public class ExportStudent
    {
        public readonly Date? BirthDate = null;
        public readonly string FirstName = null;
        public readonly EcfGender? Gender = null;
        public readonly string Id;
        public readonly string LastName = null;
        public readonly string MiddleName = null;
        public readonly string NickName = null;
        public readonly string Salutation = null;

        public ExportStudent(Configuration config, CsvTableReader csvTableReader)
        {
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Vorname"), out FirstName);
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Mittelname"), out MiddleName);
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Nachname"), out LastName);
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Geburtstag"), out BirthDate);
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Rufname"), out NickName);
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Geschlecht"), out Gender);
            csvTableReader.TryGetValue(config.GetCsvHeaderName("Anrede"), out Salutation);

            if (csvTableReader.TryGetValue(config.GetCsvHeaderName("Name"), out string name))
            {
                var csvLineParser = new CsvLineParser(' ');

                var parts = csvLineParser.Read(name);

                if (parts.Length == 2)
                {
                    FirstName = parts[0];
                    LastName = parts[1];
                }
                else if (parts.Length == 3)
                {
                    FirstName = parts[0];
                    MiddleName = parts[1];
                    LastName = parts[2];
                }
            }

            if (!csvTableReader.TryGetValue(config.GetCsvHeaderName("Id"), out Id))
            {
                Id = GenerateId();
            }
        }

        public ExportStudent(Configuration config, XlsReader xlsReader)
        {
            xlsReader.TryGetValue(config.GetXlsColumnName("Vorname"), out FirstName);
            xlsReader.TryGetValue(config.GetXlsColumnName("Mittelname"), out MiddleName);
            xlsReader.TryGetValue(config.GetXlsColumnName("Nachname"), out LastName);
            xlsReader.TryGetValue(config.GetXlsColumnName("Geburtstag"), out BirthDate);
            xlsReader.TryGetValue(config.GetXlsColumnName("Rufname"), out NickName);
            xlsReader.TryGetValue(config.GetXlsColumnName("Geschlecht"),  out Gender);
            xlsReader.TryGetValue(config.GetXlsColumnName("Anrede"), out Salutation);

            if (xlsReader.TryGetValue(config.GetXlsColumnName("Name"), out string name))
            {
                var csvLineParser = new CsvLineParser(' ');

                var parts = csvLineParser.Read(name);
                
                if (parts.Length == 2)
                {
                    FirstName = parts[0];
                    LastName = parts[1];
                }
                else if (parts.Length == 3)
                {
                    FirstName = parts[0];
                    MiddleName = parts[1];
                    LastName = parts[2];
                }
            }

            if (!xlsReader.TryGetValue(config.GetXlsColumnName("Id"), out Id))
            {
                Id = GenerateId();
            }
        }

        private string GenerateId()
        {
           var csvLineBuilder = new CsvLineBuilder();

           var keyValues = csvLineBuilder.Write(FirstName, MiddleName, LastName, BirthDate?.ToString("yyyy-MM-dd"));

           return GuidGenerator.Create(GuidGenerator.DnsNamespace, keyValues).ToString();
        }
    }
}
