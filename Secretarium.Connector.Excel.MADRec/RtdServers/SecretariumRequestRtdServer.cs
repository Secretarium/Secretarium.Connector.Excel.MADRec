using Newtonsoft.Json.Linq;
using Secretarium.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Secretarium.Excel
{
    public static class SecretariumRequestRtdServer
    {
        private static bool _isStarted;

        public const string Name = "Request";
        public const string DefaultValue = "sending...";

        private static object _locker = new object();
        private static Dictionary<XlRequest, string> XlRequestToRequestId = new Dictionary<XlRequest, string>();
        private static Dictionary<string, string> RequestIdToPayload = new Dictionary<string, string>();

        public static void EnsureStarted()
        {
            if (_isStarted) return;

            _isStarted = true;
            SecretariumFunctions.Swss.OnMessage += OnMessage;
        }

        private static void OnMessage(Byte[] data)
        {
            var message = data.GetUtf8String();
            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.DebugFormat(@"New message {0}", message);

            try
            {
                var json = JObject.Parse(message);

                var requestId = json["requestId"]?.Value<string>();
                if (string.IsNullOrEmpty(requestId))
                    return;

                if (SecretariumFunctions.Logger.IsDebugEnabled)
                    SecretariumFunctions.Logger.DebugFormat(@"Raising new message with request-id {0} {1}", requestId, message);

                var error = json["error"]?.Value<string>();
                if (!string.IsNullOrEmpty(error))
                {
                    SecretariumRtdServer.SetValue(requestId, error);
                    return;
                }

                var state = json["state"]?.Value<string>();
                if (!string.IsNullOrEmpty(state))
                {
                    SecretariumRtdServer.SetValue(requestId, state + " @" + DateTime.Now);
                    return;
                }

                var result = json["result"];
                if (result != null)
                {
                    SecretariumRtdServer.SetValue(requestId, result.ToJson(true));
                    return;
                }
            }
            catch (Exception ex)
            {
                SecretariumFunctions.Logger.Error("Can't parse", ex);
            }
        }

        public static bool OnConnectData(ref Array parameters, out string requestId, out string error)
        {
            requestId = null;
            error = null;

            if (parameters.Length != 2)
            {
                error = "Invalid parameters";
                return false;
            }

            if (SecretariumFunctions.Swss.State != SwssConnector.ConnectionState.SecureConnectionEstablished)
            {
                error = "Please connect first";
                return false;
            }

            requestId = parameters.GetValue(1).ToString();
            if(!RequestIdToPayload.TryGetValue(requestId, out string payload))
            {
                error = "Invalid request id";
                return false;
            }

            SecretariumFunctions.Swss.Send(payload);
            
            return true;
        }

        public static bool TryGet(XlRequest xlReq, out string requestId)
        {
            lock (_locker)
            {
                return XlRequestToRequestId.TryGetValue(xlReq, out requestId);
            }
        }

        public static void Add(XlRequest xlReq, string requestId, string payload)
        {
            lock (_locker)
            {
                XlRequestToRequestId[xlReq] = requestId;
                RequestIdToPayload[requestId] = payload;
            }
        }

        /// <summary>
        /// Everytime the value of the calling cell is updated by the Rtd server, Excel calls the UDF again.
        /// We need a way to identify a Request from its parameters to avoid querying the platform again for technical reasons.
        /// </summary>
        public class XlRequest : IEquatable<XlRequest>
        {
            public readonly int HashCode;

            public readonly string[] Args;

            public XlRequest(params string[] args)
            {
                Args = args?.Select(x => x.ToUpper()).ToArray();
                HashCode = args == null ? 0 : ComputeHashCode();
            }

            private int ComputeHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    for(int i = 0; i < Args.Length; i++)
                    {
                        hash = hash * 23 + (Args[i] == null ? 0 : Args[i].GetHashCode());
                    }
                    return hash;
                }
            }

            public bool Equals(XlRequest other)
            {
                return other != null && HashCode == other.HashCode && 
                    ((Args == null && other.Args == null) || (Args != null && other.Args != null && Args.SequenceEqual(other.Args)));
            }

            public override bool Equals(object other)
            {
                return Equals(other as XlRequest);
            }

            public override int GetHashCode()
            {
                return HashCode;
            }
        }
    }
}
