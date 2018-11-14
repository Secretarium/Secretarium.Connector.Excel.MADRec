using ExcelDna.Integration;
using log4net;
using Secretarium.Helpers;
using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace Secretarium.Excel
{
    public partial class SecretariumFunctions : IExcelAddIn
    {
        internal static SecureConnectionProtocol Scp = new SecureConnectionProtocol();
        private static ScpConfig _config;

        internal static readonly ILog Logger = LogManager.GetLogger("SecretariumExcelAddin");

        public void AutoOpen()
        {
            Logger.InfoFormat(@"Starting");
            
            // Config
            _config = new ScpConfig
            {
                client = new ScpConfig.ClientConfig
                {
                    proofOfWorkMaxDifficulty = 18
                },
                secretarium = new ScpConfig.SecretariumConfig
                {
                    endPoint = "wss://ovh1.node.secretarium.org:443/",
                    knownPubKey = "rliD_CISqPEeYKbWYdwa-L-8oytAPvdGmbLC0KdvsH-OVMraarm1eo-q4fte0cWJ7-kmsq8wekFIJK0a83_yCg=="
                }
            };
            Scp.Init(_config);

            Logger.InfoFormat(@"Starting with endpoint:{0} and trustedley:{1}", _config.secretarium.endPoint, _config.secretarium.knownPubKey);
        }

        public void AutoClose()
        {
            Logger.InfoFormat(@"Closing");
            Scp.Dispose();
            Logger.InfoFormat(@"Closed");
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

        private static object SendRt<T>(SecretariumRequestRtdServer.XlRequest xlReq, string DCApp, string function, T args) where T : class
        {
            if (Scp.State.IsClosed())
                return "Please connect first";

            var request = new Request<T>(DCApp, function, args);

            SecretariumRequestRtdServer.Add(xlReq, request.requestId, request.ToJson());

            SecretariumRequestRtdServer.EnsureStarted();
            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumRequestRtdServer.Name, request.requestId);
        }


        [ExcelFunction("Sets your identity from a Secretarium key", Category = "Secretarium", Name = "Secretarium.SetIdentityFromSecKey")]
        public static string SetIdentityFromSecKey([ExcelArgument("Path to Secretarium key file")] string secFile, [ExcelArgument("Key password")] string password)
        {
            if (string.IsNullOrEmpty(secFile))
                return "Missing Secretarium key file";

            if (string.IsNullOrEmpty(password))
                return "Missing password";

            try
            {
                var cfg = JsonHelper.DeserializeJsonFromFileAs<ScpConfig.KeyConfig>(secFile);
                cfg.password = password;
                
                if (cfg.TryGetECDsaKey(out ECDsaCng key))
                    Scp.Set(key);
                else
                    return "Could not load your identity";
            }
            catch (Exception ex)
            {
                return "Could not load your identity " + ex.Message;
            }

            Logger.InfoFormat(@"New identity set:{0}", Scp.PublicKey);

            return Scp.PublicKey;
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
                    Scp.Set(key);
                else
                    return "Could not load your identity";
            }
            catch(Exception ex)
            {
                return "Could not load your identity " + ex.Message;
            }
            
            Logger.InfoFormat(@"New identity set:{0}", Scp.PublicKey);

            return Scp.PublicKey;
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
                    Scp.Set(key);
                else
                    return "Could not load your identity";
            }
            catch (Exception ex)
            {
                return "Could not load your identity " + ex.Message;
            }

            Logger.InfoFormat(@"New identity set:{0}", Scp.PublicKey);

            return Scp.PublicKey;
        }


        [ExcelFunction("Override Secretarium trusted key", Category = "Secretarium", Name = "Secretarium.SetTrustedKey")]
        public static string OverrideSecretariumTrustedKey([ExcelArgument("The trusted Secretarium public key")] string trustedKey)
        {
            if (string.IsNullOrEmpty(trustedKey))
                return "Missing trusted key";

            Logger.InfoFormat(@"New trusted key:{0}", trustedKey);
            _config.secretarium.knownPubKey = trustedKey;

            Scp.Disconnect();

            return "Trusted key overridden";
        }

        [ExcelFunction("Override Secretarium endpoint", Category = "Secretarium", Name = "Secretarium.SetGateway")]
        public static string OverrideSecretariumEndpoint([ExcelArgument("The Secretarium gateway endpoint")] string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint))
                return "Missing endpoint";

            Logger.InfoFormat(@"New endpoint:{0}", endpoint);
            _config.secretarium.endPoint = endpoint;

            Scp.Disconnect();

            return "Endpoint overridden to " + endpoint;
        }
        

        [ExcelFunction("Connect to Secretarium", Category = "Secretarium", Name = "Secretarium.Rt.Connect")]
        public static object ConnectRt()
        {
            if (Scp.PublicKey == null)
                return "Please set your identity first";

            if (Scp.State.IsClosed())
                new Task(() => Scp.Connect()).Start();

            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumConnectionStateRtdServer.Name);
        }

        [ExcelFunction("Connect to Secretarium", Category = "Secretarium", Name = "Secretarium.Rt.ConnectionState")]
        public static object ConnectionState()
        {
            SecretariumConnectionStateRtdServer.EnsureStarted();
            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, SecretariumConnectionStateRtdServer.Name);
        }


        [ExcelFunction("View a message after AES256 encryption, base64 encoded", Category = "Secretarium", Name = "Secretarium.Encrypt")]
        public static string Encrypt([ExcelArgument("A message")] string message)
        {
            if (Scp.SymmetricKey == null)
                return "No key defined yet, please connect first";

            var ivOffset = ByteHelper.GetRandom(16);
            return message.ToBytes().AesCtrEncrypt(Scp.SymmetricKey, ivOffset).ToBase64String();
        }

        [ExcelFunction("View value after Sha256 hashing", Category = "Secretarium", Name = "Secretarium.HashSha256")]
        public static string HashSha256( [ExcelArgument("The value")] object value)
        {
            if (value == null)
                return "Invalid value";

            if (!(value is string || value is int || value is double || value is bool))
                return "Invalid value type " + value.GetType();

            if (value is string)
                return ((string)value).HashSha256().ToBase64String(false);
            else if (value is int)
                return ((int)value).HashSha256().ToBase64String(false);
            else if (value is double)
                return ((double)value).HashSha256().ToBase64String(false);
            return ((bool)value).HashSha256().ToBase64String(false);
        }


        [ExcelFunction("Send to Secretarium", Category = "Secretarium", Name = "Secretarium.RtSend")]
        public static object SendRt([ExcelArgument("The DCApp name")] string DCApp, [ExcelArgument("The function name")] string function, [ExcelArgument("The args as JSON")] string argsJson)
        {
            if (Scp.State.IsClosed())
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
    }
}