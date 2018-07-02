using ExcelDna.Integration;
using log4net;
using Secretarium.Helpers;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace Secretarium.Excel
{
    public class SecretariumFunctions : IExcelAddIn
    {
        internal static SwssConnector Swss = new SwssConnector();
        private static SwssConfig _config;

        internal static readonly ILog Logger = LogManager.GetLogger("SecretariumExcelAddin");

        public void AutoOpen()
        {
            Logger.InfoFormat(@"Starting");
            
            // Config
            _config = new SwssConfig
            {
                client = new SwssConfig.ClientConfig
                {
                    proofOfWorkMaxDifficulty = 18
                },
                secretarium = new SwssConfig.SecretariumConfig
                {
                    endPoint = "wss://swss.secretarium.org:5424/",
                    knownPubKey = "rliD_CISqPEeYKbWYdwa-L-8oytAPvdGmbLC0KdvsH-OVMraarm1eo-q4fte0cWJ7-kmsq8wekFIJK0a83_yCg=="
                }
            };
            Swss.Init(_config);

            Logger.InfoFormat(@"Starting with endpoint:{0} and trustedley:{1}", _config.secretarium.endPoint, _config.secretarium.knownPubKey);
        }

        public void AutoClose()
        {
            Logger.InfoFormat(@"Closing");
            Swss.Dispose();
            Logger.InfoFormat(@"Closed");
        }


        [ExcelFunction("Sets your identity from a X509 certificate", Category = "Secretarium", Name = "Secretarium.SetIdentityFromX509")]
        public static string SetIdentityFromX509 ([ExcelArgument("Path to X509 pfx certificate")] string pfxFile, [ExcelArgument("X509 password")] string password)
        {
            if (string.IsNullOrEmpty(pfxFile))
                return "Missing pfx file";

            if (string.IsNullOrEmpty(password))
                return "Missing password";

            try
            {
                var x509 = X509Helper.LoadX509FromFile(pfxFile, password);
                if (x509.GetECDsaPrivateKey() is ECDsaCng key && key.HashAlgorithm == CngAlgorithm.Sha256 && key.KeySize == 256)
                    Swss.Set(key);
                else
                    return "Could not load your identity";
            }
            catch(Exception ex)
            {
                return "Could not load your identity " + ex.Message;
            }
            
            Logger.InfoFormat(@"New identity set:{0}", Swss.PublicKey);

            return Swss.PublicKey;
        }

        [ExcelFunction("Sets your identity", Category = "Secretarium", Name = "Secretarium.SetIdentity")]
        public static string SetIdentity([ExcelArgument("Base 64 encoded public key")] string publicKey, [ExcelArgument("Base 64 encoded private key")] string privateKey)
        {
            if (string.IsNullOrEmpty(publicKey))
                return "Missing public key";

            if (string.IsNullOrEmpty(privateKey))
                return "Missing private key";

            try
            {
                var key = ECDsaHelper.Import(publicKey.FromBase64String(), privateKey.FromBase64String());
                if (key != null && key.HashAlgorithm == CngAlgorithm.Sha256 && key.KeySize == 256)
                    Swss.Set(key);
                else
                    return "Could not load your identity";
            }
            catch (Exception ex)
            {
                return "Could not load your identity " + ex.Message;
            }

            Logger.InfoFormat(@"New identity set:{0}", Swss.PublicKey);

            return Swss.PublicKey;
        }

        [ExcelFunction("Override Secretarium trusted key", Category = "Secretarium", Name = "Secretarium.SetTrustedKey")]
        public static string OverrideSecretariumTrustedKey([ExcelArgument("The trusted Secretarium public key")] string trustedKey)
        {
            if (string.IsNullOrEmpty(trustedKey))
                return "Missing trusted key";

            Logger.InfoFormat(@"New trusted key:{0}", trustedKey);
            _config.secretarium.knownPubKey = trustedKey;

            Swss.Disconnect();

            return "Trusted key overridden";
        }

        [ExcelFunction("Override Secretarium endpoint", Category = "Secretarium", Name = "Secretarium.Demo.SetEndpoint")]
        public static string OverrideSecretariumEndpoint([ExcelArgument("The Secretarium endpoint")] string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint))
                return "Missing endpoint";

            Logger.InfoFormat(@"New endpoint:{0}", endpoint);
            _config.secretarium.endPoint = endpoint;

            Swss.Disconnect();

            return "Endpoint overridden to " + endpoint;
        }
        

        [ExcelFunction("Connect to Secretarium", Category = "Secretarium", Name = "Secretarium.Rt.Connect")]
        public static object ConnectRt()
        {
            if (Swss.PublicKey == null)
                return "Please set your identity first";

            if (Swss.State.IsClosed())
                new Task(() => Swss.Connect()).Start();

            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumConnectionStateRtdServer.Name);
        }

        [ExcelFunction("Connect to Secretarium", Category = "Secretarium", Name = "Secretarium.Rt.ConnectionState")]
        public static object ConnectionState()
        {
            SecretariumConnectionStateRtdServer.EnsureStarted();
            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumConnectionStateRtdServer.Name);
        }

        [ExcelFunction("View a message after AES256 encryption, base64 encoded", Category = "Secretarium", Name = "Secretarium.Demo.Encrypt")]
        public static string Encrypt([ExcelArgument("A message")] string message)
        {
            if (Swss.SymmetricKey == null)
                return "No key defined yet, please connect first";

            var ivOffset = ByteHelper.GetRandom(16);
            return message.ToBytes().AesCtrEncrypt(Swss.SymmetricKey, ivOffset).ToBase64String();
        }

        private static bool GetRequest(string DCApp, string function, string argsJson, out Request request, out string error)
        {
            request = null;
            error = null;

            if (string.IsNullOrWhiteSpace(DCApp))
            {
                error = "Invalid DCApp";
                return false;
            }

            if (string.IsNullOrWhiteSpace(function))
            {
                error = "Invalid function";
                return false;
            }

            switch (DCApp.ToLower())
            {
                case "madrec":
                    switch (function.ToLower())
                    {
                        case "put":
                            if (!argsJson.TryDeserializeJsonAs(out MADRecPutLEIArgs mpla))
                            {
                                error = "Invalid args";
                                return false;
                            }
                            request = new Request<MADRecPutLEIArgs>(DCApp, function, mpla);
                            break;
                        case "get":
                            if (!argsJson.TryDeserializeJsonAs(out MADRecGetLEIArgs mgla))
                            {
                                error = "Invalid args";
                                return false;
                            }
                            request = new Request<MADRecGetLEIArgs>(DCApp, function, mgla);
                            break;
                        default:
                            error = "Invalid function";
                            return false;
                    }
                    break;
                case "dcappfortesting":
                    switch (function.ToLower())
                    {
                        case "sum":
                        case "avg":
                            if (!argsJson.TryDeserializeJsonAs(out double[] ad))
                            {
                                error = "Invalid args";
                                return false;
                            }
                            request = new Request<double[]>(DCApp, function, ad);
                            break;
                        default:
                            error = "Invalid function";
                            return false;
                    }
                    break;
                default:
                    error = "Invalid DCApp";
                    return false;
            }

            return true;
        }

        [ExcelFunction("Send to Secretarium", Category = "Secretarium", Name = "Secretarium.RtSend")]
        public static object SendRt([ExcelArgument("The DCApp name")] string DCApp, [ExcelArgument("The function name")] string function, [ExcelArgument("The args as JSON")] string argsJson)
        {
            if (Swss.State.IsClosed())
                return "Please connect first";

            if (string.IsNullOrWhiteSpace(argsJson) || argsJson.Length > 255)
                return "Invalid args (max 255 chars)";

            // Excel comes back to the UDF when its value updates
            var xlRequest = new SecretariumRequestRtdServer.XlRequest(DCApp, function, argsJson);
            if (SecretariumRequestRtdServer.TryGet(xlRequest, out string requestId))
                return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, requestId);

            switch (DCApp.ToLower())
            {
                case "madrec":
                    switch (function.ToLower())
                    {
                        case "put":
                            if (!argsJson.TryDeserializeJsonAs(out MADRecPutLEIArgs mpla))
                                return "Invalid args";
                            return SendRt(xlRequest, DCApp, function, mpla);
                        case "get":
                            if (!argsJson.TryDeserializeJsonAs(out MADRecGetLEIArgs mgla))
                                return "Invalid args";
                            return SendRt(xlRequest, DCApp, function, mgla);
                        default:
                            return "Invalid function";
                    }
                case "dcappfortesting":
                    switch (function.ToLower())
                    {
                        case "sum":
                        case "avg":
                            if (!argsJson.TryDeserializeJsonAs(out double[] ad))
                                return "Invalid args";
                            return SendRt(xlRequest, DCApp, function, ad);
                        default:
                            return "Invalid function";
                    }
                default:
                    return "Invalid DCApp";
            }
        }

        private static object SendRt<T>(SecretariumRequestRtdServer.XlRequest xlReq, string DCApp, string function, T args) where T : class
        {
            if (Swss.State.IsClosed())
                return "Please connect first";

            var request = new Request<T>(DCApp, function, args);

            SecretariumRequestRtdServer.Add(xlReq, request.requestId, request.ToJson());

            SecretariumRequestRtdServer.EnsureStarted();
            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, request.requestId);
        }


        [ExcelFunction("Sends to the MADRec DCApp", Category = "Secretarium", Name = "MADRec.RtPut")]
        public static object MadrecPut([ExcelArgument("The LEI")] string LEI, [ExcelArgument("The args as JSON")] string argsJson)
        {
            if (Swss.State.IsClosed())
                return "Please connect first";

            if (string.IsNullOrWhiteSpace(LEI))
                return "Invalid LEI";

            if (string.IsNullOrWhiteSpace(argsJson))
                return "Invalid args";

            // Excel comes back to the UDF when its value updates
            var xlRequest = new SecretariumRequestRtdServer.XlRequest("MADRec", "put", LEI, argsJson);
            if(SecretariumRequestRtdServer.TryGet(xlRequest, out string requestId))
                return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, requestId);

            if (!argsJson.TryDeserializeJsonAs(out List<MADRecPutValues> mpv))
                return "Invalid args";
            
            return SendRt(xlRequest, "MADRec", "put", new MADRecPutLEIArgs { LEI = LEI, values = mpv });
        }

        [ExcelFunction("Gets from the MADRec DCApp", Category = "Secretarium", Name = "MADRec.RtGet")]
        public static object MadrecGet([ExcelArgument("The LEI")] string LEI)
        {
            if (Swss.State.IsClosed())
                return "Please connect first";

            if (string.IsNullOrWhiteSpace(LEI))
                return "Invalid LEI";

            // Excel comes back to the UDF when its value updates
            var xlRequest = new SecretariumRequestRtdServer.XlRequest("MADRec", "get", LEI);
            if (SecretariumRequestRtdServer.TryGet(xlRequest, out string requestId))
                return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, requestId);

            return SendRt(xlRequest, "MADRec", "get", new MADRecGetLEIArgs { LEI = LEI });
        }

        [ExcelFunction("Formatting helper for the MADRec DCApp", Category = "Secretarium", Name = "MADRec.Format.Pair")]
        public static string MadrecFormatPair([ExcelArgument("The field name")] string name, [ExcelArgument("The field value")] object value)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Invalid name";

            if (value == null)
                return "Invalid value";

            if (value is string || value is int || value is double || value is bool)
                return new MADRecContrib { name = name, value = value }.ToJson(true);

            return "Invalid value type " + value.GetType();
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
            o[7, 0] = extract.total == 1 ? 
                "no match" : extract.split.Length == 1 ? 
                "full consensus" : extract.split.Length < extract.total ?
                "partial consensus" + (extract.group == extract.split[0] ? "" : " (out)") : "no consensus";
            for (int i = 0; i < extract.split.Length; i++)
            {
                o[8 + i, 0] = extract.split[i];
            }
            return o;
        }
    }
}