using ExcelDna.Integration;
using Secretarium.Helpers;
using System;
using System.Collections.Generic;

namespace Secretarium.Excel
{
    public partial class SecretariumFunctions
    {
        [ExcelFunction("Sends to the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.Verify")]
        public static object MadrecVerify([ExcelArgument("The field name")] string field, [ExcelArgument("The value")] object value = null)
        {
            if (string.IsNullOrWhiteSpace(field) || !MADRecFormats.MADREC_FIELDS.Contains(field))
                return "Invalid field name";

            if (value == null || value == ExcelEmpty.Value || (value is string && string.IsNullOrEmpty(value as string)))
            {
                if (field == "LEI")
                    return "Invalid value";
                else
                    return "";
            }

            if (MADRecFormats.Verify(field, ref value, out string error))
                return value;

            return "ERROR:" + error;
        }


        [ExcelFunction("Sends to the MADRec DCApp", Category = "Secretarium", Name = "MADRec.RtPut")]
        public static object MadrecPut([ExcelArgument("The LEI")] string LEI, [ExcelArgument("The args as JSON")] string argsJson, [ExcelArgument("Readiness status")] bool hashed)
        {
            if (Scp.State.IsClosed())
                return "Please connect first";

            if (string.IsNullOrWhiteSpace(LEI))
                return "Invalid LEI";

            if (string.IsNullOrWhiteSpace(argsJson))
                return "Invalid args";

            // Excel comes back to the UDF when its value updates
            var xlRequest = new SecretariumRequestRtdServer.XlRequest("madrec", "put", LEI, argsJson);
            if(SecretariumRequestRtdServer.TryGet(xlRequest, out string requestId))
                return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, requestId);

            if (!argsJson.TryDeserializeJsonAs(out List<MADRecPutValues> mpv))
                return "Invalid args";
            
            return SendRt(xlRequest, "madrec", "put", new MADRecPutLEIArgs { LEI = LEI, values = mpv, hashed = hashed });
        }

        [ExcelFunction("Gets from the MADRec DCApp", Category = "Secretarium", Name = "MADRec.RtGet")]
        public static object MadrecGet([ExcelArgument("The LEI")] string LEI, [ExcelArgument("Subscribe")] bool subscribe = true)
        {
            if (Scp.State.IsClosed())
                return "Please connect first";

            if (string.IsNullOrWhiteSpace(LEI))
                return "Invalid LEI";

            // Excel comes back to the UDF when its value updates
            var xlRequest = new SecretariumRequestRtdServer.XlRequest("madrec", "get", LEI);
            if (SecretariumRequestRtdServer.TryGet(xlRequest, out string requestId))
                return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, requestId);

            return SendRt(xlRequest, "madrec", "get", new MADRecGetLEIArgs { LEI = LEI, subscribe = subscribe });
        }


        [ExcelFunction("Formatting helper for the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.Pair")]
        public static string MadrecFormatPair([ExcelArgument("The field name")] string name, [ExcelArgument("The field value")] object value, [ExcelArgument("hash")] bool hash = true)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Invalid name";

            if (value == null)
                return "Invalid value";
            
            if (!(value is string || value is int || value is double || value is bool))
                return "Invalid value type " + value.GetType();

            if (hash)
            {
                if (value is string) value = ((string)value).HashSha256().ToBase64String(false);
                else if (value is int) value = ((int)value).HashSha256().ToBase64String(false);
                else if (value is double) value = ((double)value).HashSha256().ToBase64String(false);
                else value = ((bool)value).HashSha256().ToBase64String(false);
            }

            return new MADRecContrib { name = name, value = value }.ToJson(true);
        }

        [ExcelFunction("Formatting helper for the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.Pairs")]
        public static string MadrecFormatPairs(
            string name1, object value1, string name2, object value2, string name3, object value3, string name4, object value4, string name5, object value5,
            string name6, object value6, string name7, object value7, string name8, object value8, string name9, object value9, string name10, object value10)
        {
            var allNames = new[] { name1, name2, name3, name4, name5, name6, name7, name8, name9, name10 };
            var allValues = new[] { value1, value2, value3, value4, value5, value6, value7, value8, value9, value10 };

            var contribs = new List<MADRecContrib>();

            for (var i = 0; i < 10; i++)
            {
                if (string.IsNullOrWhiteSpace(allNames[i]))
                    return contribs.ToJson(true);

                if (allValues[i] == null)
                    return contribs.ToJson(true);

                if (allValues[i] is string || allValues[i] is int || allValues[i] is double || allValues[i] is bool)
                    contribs.Add(new MADRecContrib { name = allNames[i], value = allValues[i] });

                else
                    return "Invalid value #" + (i + 1);
            }

            return contribs.ToJson(true);
        }

        [ExcelFunction("Formatting helper for the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.Combine")]
        public static string MadrecFormatCombine(
            string arg1, string arg2, string arg3, string arg4, string arg5, string arg6, string arg7, string arg8, string arg9, string arg10,
            string arg11, string arg12, string arg13, string arg14, string arg15, string arg16, string arg17, string arg18, string arg19, string arg20)
        {
            var all = new[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 };

            var contribs = new List<MADRecPutValues>();
            
            for(var i = 0; i< 10; i++)
            {
                if (string.IsNullOrWhiteSpace(all[i]))
                    return contribs.ToJson(true);

                if (all[i].TryDeserializeJsonAs(out MADRecPutValues mc))
                    contribs.Add(mc);
                else if (all[i].TryDeserializeJsonAs(out List<MADRecPutValues> mcarr))
                    contribs.AddRange(mcarr);
                else
                    return "Invalid arg #" + (i + 1);
            }

            return contribs.ToJson(true);
        }

        [ExcelFunction("Formatting helper for the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.SubContrib")]
        public static string MadrecFormatSubContrib([ExcelArgument("The field name")] string name, [ExcelArgument("The combined values as JSON")] string valuesJson)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Invalid name";

            if (string.IsNullOrWhiteSpace(valuesJson))
                return "Invalid values";

            if (!valuesJson.TryDeserializeJsonAs(out List<MADRecPutValues> madrecPutValues))
                return "Invalid values";

            return new MADRecPutValues { name = name, values = madrecPutValues }.ToJson(true);
        }
        
        [ExcelFunction("Sends to the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.Contrib")]
        public static string MadrecFormatContrib([ExcelArgument("The LEI")] string LEI, [ExcelArgument("The args as JSON")] string argsJson)
        {
            if (string.IsNullOrWhiteSpace(LEI))
                return "Invalid LEI";

            if (string.IsNullOrWhiteSpace(argsJson))
                return "Invalid args";

            if (!argsJson.TryDeserializeJsonAs(out List<MADRecPutValues> mpv))
                return "Invalid args";

            return new MADRecPutLEIArgs { LEI = LEI, values = mpv }.ToJson(true);
        }


        private static string GetConsensusState(MADRecFieldResult report)
        {
            return report.total == 1 ?
                "no match" : report.split.Length == 1 ?
                "full consensus" : report.split.Length < report.total ?
                "split consensus with " + (report.group == report.split[0] ? "majority" : "minority") : "no consensus";
        }

        [ExcelFunction("Extract one field report from a MADRec result", Category = "Secretarium", Name = "MADRec.Extract.ToPieChartData")]
        public static object[,] MadrecToPieChartData([ExcelArgument("The MADRec report")] string report, [ExcelArgument("The field name, use '/' to target subfields")] string field)
        {
            if (string.IsNullOrWhiteSpace(field))
                return new object[,] { { "Invalid field" }, { "" }, { "" }, { 0 }, { 0 }, { "" }, { "" }, { "" }, { 0 }, { 0 } };

            if (string.IsNullOrWhiteSpace(report) || !report.TryDeserializeJsonAs(out MADRecResult mr))
                return new object[,] { { "Invalid report" }, { "" }, { "" }, { 0 }, { 0 }, { "" }, { "" }, { "" }, { 0 }, { 0 } };

            var extract = mr.FindReport(field.Split('/'));
            if (extract == null)
                return new object[,] { { "field not found" }, { "" }, { "" }, { 0 }, { 0 }, { "" }, { "" }, { "" }, { 0 }, { 0 } };

            Array.Sort(extract.split);
            Array.Reverse(extract.split);
            var o = new object[8 + extract.split.Length, 1];
            o[0, 0] = mr.LEI + " - " + field;
            o[1, 0] = extract.name; // field;
            o[2, 0] = extract.contribution;
            o[3, 0] = extract.total;
            o[4, 0] = extract.group;
            o[5, 0] = extract.split.ToJson();
            o[6, 0] =
                "Contrib: " + extract.contribution + "\n" +
                "Total: " + extract.total + "\n" +
                "Group: " + extract.group + "\n" +
                "Split: " + o[5, 0];
            o[7, 0] = GetConsensusState(extract);
            for (int i = 0; i < extract.split.Length; i++)
            {
                o[8 + i, 0] = extract.split[i];
            }
            return o;
        }

        [ExcelFunction("Extract one field report from a MADRec result", Category = "Secretarium", Name = "MADRec.Extract.ConsensusState")]
        public static string MadrecConsensusState([ExcelArgument("The MADRec report")] string report, [ExcelArgument("The field name, use '/' to target subfields")] string field)
        {
            if (string.IsNullOrWhiteSpace(field))
                return "Invalid field";

            if (string.IsNullOrWhiteSpace(report) || !report.TryDeserializeJsonAs(out MADRecResult mr))
                return "Invalid report";

            var extract = mr.FindReport(field.Split('/'));
            if (extract == null)
                return "field not found";

            Array.Sort(extract.split);
            Array.Reverse(extract.split);            
            return GetConsensusState(extract);
        }
    }
}