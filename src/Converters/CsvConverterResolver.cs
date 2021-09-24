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

using Enbrea.Csv;
using Enbrea.Ecf;
using System;

namespace Ecf.Excel
{
    /// <summary>
    /// Own implementation of an <see cref="ICsvConverterResolver"/>
    /// </summary>
    public class CsvConverterResolver : CsvDefaultConverterResolver
    {
        protected override void RegisterDefaultConverters()
        {
            AddConverter(typeof(bool), new CsvBooleanConverter());
            AddConverter(typeof(byte), new CsvByteConverter());
            AddConverter(typeof(char), new CsvCharConverter());
            AddConverter(typeof(Date), new CsvDateConverter());
            AddConverter(typeof(decimal), new CsvDecimalConverter());
            AddConverter(typeof(double), new CsvDoubleConverter());
            AddConverter(typeof(Guid), new CsvGuidConverter());
            AddConverter(typeof(int), new CsvInt32Converter());
            AddConverter(typeof(long), new CsvInt64Converter());
            AddConverter(typeof(sbyte), new CsvSByteConverter());
            AddConverter(typeof(short), new CsvInt16Converter());
            AddConverter(typeof(string), new CsvStringConverter());
            AddConverter(typeof(uint), new CsvUInt32Converter());
            AddConverter(typeof(uint), new CsvUInt32Converter());
            AddConverter(typeof(ulong), new CsvUInt64Converter());
            AddConverter(typeof(ulong), new CsvUInt64Converter());
            AddConverter(typeof(Uri), new CsvUriConverter());
            AddConverter(typeof(ushort), new CsvUInt16Converter());
            AddConverter(typeof(EcfGender), new CsvGenderConverter());
        }
    }
}