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

namespace Ecf.Excel
{
    /// <summary>
    /// Default implementation of an <see cref="ICsvConverterResolver"/>
    /// </summary>
    public class CsvConverterResolver : CsvDefaultConverterResolver
    {
        protected override void RegisterDefaultConverters()
        {
            AddConverter<bool?>(new CsvBooleanConverter());
            AddConverter<bool>(new CsvBooleanConverter());
            AddConverter<byte?>(new CsvByteConverter());
            AddConverter<byte>(new CsvByteConverter());
            AddConverter<char?>(new CsvCharConverter());
            AddConverter<char>(new CsvCharConverter());
            AddConverter<Date?>(new CsvDateConverter());
            AddConverter<Date>(new CsvDateConverter());
            AddConverter<decimal?>(new CsvDecimalConverter());
            AddConverter<decimal>(new CsvDecimalConverter());
            AddConverter<double?>(new CsvDoubleConverter());
            AddConverter<double>(new CsvDoubleConverter());
            AddConverter<EcfGender?>(new CsvGenderConverter());
            AddConverter<EcfGender>(new CsvGenderConverter());
            AddConverter<Guid?>(new CsvGuidConverter());
            AddConverter<Guid>(new CsvGuidConverter());
            AddConverter<int?>(new CsvInt32Converter());
            AddConverter<int>(new CsvInt32Converter());
            AddConverter<long?>(new CsvInt64Converter());
            AddConverter<long>(new CsvInt64Converter());
            AddConverter<sbyte?>(new CsvSByteConverter());
            AddConverter<sbyte>(new CsvSByteConverter());
            AddConverter<short?>(new CsvInt16Converter());
            AddConverter<short>(new CsvInt16Converter());
            AddConverter<string>(new CsvStringConverter());
            AddConverter<uint?>(new CsvUInt32Converter());
            AddConverter<uint>(new CsvUInt32Converter());
            AddConverter<ulong?>(new CsvUInt64Converter());
            AddConverter<ulong>(new CsvUInt64Converter());
            AddConverter<ushort?>(new CsvUInt16Converter());
            AddConverter<ushort>(new CsvUInt16Converter());
        }
    }
}