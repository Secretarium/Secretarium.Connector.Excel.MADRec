using ExcelDna.Integration;
using log4net;
using Newtonsoft.Json.Linq;
using Secretarium.Helpers;
using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace Secretarium.Excel
{
    public class SecretariumFunctions : IExcelAddIn
    {
        private static ScpConfig _config;

        internal static readonly SecureConnectionProtocol Scp = new SecureConnectionProtocol();
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


        [ExcelFunction("View a message after AES256 encryption, base64 encoded", Category = "Secretarium", Name = "Secretarium.Encrypt")]
        public static string Encrypt([ExcelArgument("A message")] string message)
        {
            if (Scp.SymmetricKey == null)
                return "No key defined yet, please connect first";

            var ivOffset = ByteHelper.GetRandom(16);
            return message.ToBytes().AesCtrEncrypt(Scp.SymmetricKey, ivOffset).ToBase64String();
        }

        [ExcelFunction("Parses and formats a Json", Category = "Secretarium", Name = "Secretarium.FormatJson")]
        public static string FormatJson([ExcelArgument("A json")] string json)
        {
            try
            {
                return JObject.Parse(json).ToString();
            }
            catch(Exception e)
            {
                return "Error: can't parse input. " + e.Message;
            }
        }


        [ExcelFunction("Connect to Secretarium", Category = "Secretarium", Name = "Secretarium.RtConnect")]
        public static object RtConnect()
        {
            if (Scp.PublicKey == null)
                return "Please set your identity first";

            if (Scp.State.IsClosed())
                new Task(() => Scp.Connect()).Start();

            return SecretariumRtdServer.GetRtdConnectionState();
        }

        [ExcelFunction("Connect to Secretarium", Category = "Secretarium", Name = "Secretarium.RtConnectionState")]
        public static object RtConnectionState()
        {
            return SecretariumRtdServer.GetRtdConnectionState();
        }

        [ExcelFunction("Send to Secretarium", Category = "Secretarium", Name = "Secretarium.RtSend")]
        public static object RtRequest([ExcelArgument("The DCApp name")] string DCApp, [ExcelArgument("The function name")] string function, [ExcelArgument("The args as JSON")] string argsJson)
        {
            if (Scp.State.IsClosed())
                return "Please connect first";

            return SecretariumRtdServer.GetRtdRequest(DCApp, function, argsJson);
        }
    }
}