using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Newtonsoft.Json.Linq;
using Secretarium.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Secretarium.Excel
{
    public class SecretariumRtdServer : IRtdServer
    {
        public class RtdTypes
        {
            public const string ConnectionState = "ConnectionState";
            public const string Request = "Request";
        }

        private const string DefaultValue = "loading...";
        private bool _timerStarted = false;
        private IRTDUpdateEvent _callback;
        private System.Windows.Forms.Timer _timer;
        private static object _updatesLocker = new object();
        private static Dictionary<string, object> _updates = new Dictionary<string, object>();
        private static object _mappingsLocker = new object();
        private Dictionary<string, int> _requesIdToTopics;
        private Dictionary<int, string> _topicsToRequestId;
        private static Dictionary<XlRequest, string> _xlRequestToRequestId;
        private static Dictionary<string, string> _requestIdToPayload;

        internal static object GetRtdConnectionState()
        {
            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, RtdTypes.ConnectionState);
        }

        internal static object GetRtdRequest(string DCApp, string function, string argsJson)
        {
            var xlRequest = new XlRequest(DCApp, function, argsJson);

            if (_xlRequestToRequestId.TryGetValue(xlRequest, out string requestId))
                return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, RtdTypes.Request, requestId);

            var request = new Request(DCApp, function, argsJson);

            _xlRequestToRequestId[xlRequest] = request.requestId;
            _requestIdToPayload[request.requestId] = request.ToJson();

            return XlCall.RTD("Secretarium.Excel.SecretariumRtdServer", null, RtdTypes.Request, request.requestId);
        }

        #region implement IRtdServer

        public virtual int ServerStart(IRTDUpdateEvent callback)
        {
            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.Debug(@"RequestRtdServer starting");

            _callback = callback;
            _requesIdToTopics = new Dictionary<string, int>();
            _topicsToRequestId = new Dictionary<int, string>();
            _xlRequestToRequestId = new Dictionary<XlRequest, string>();
            _requestIdToPayload = new Dictionary<string, string>();

            // This Windows.Forms Single threaded (UI thread) timer must be started in this correct COM appartment
            _timer = new System.Windows.Forms.Timer { Interval = 1000 };
            _timer.Tick += (object sender, EventArgs args) =>
            {
                _timer.Stop();

                lock (_updatesLocker)
                {
                    CleanUpUpdates();

                    if (_updates.Count > 0)
                        _callback.UpdateNotify();
                }

                _timer.Start();
            };

            SecretariumFunctions.Scp.OnMessage += OnMessage;
            SecretariumFunctions.Scp.OnStateChange += x =>
            {
                SetValue(RtdTypes.ConnectionState, SecretariumFunctions.Scp.State.ToString());
            };

            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.Debug(@"RequestRtdServer started");

            return 1; // 1 means OK
        }

        public object ConnectData(int topicId, ref Array parameters, ref bool newValues)
        {
            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.DebugFormat(@"RequestRtdServer topic connecting {0}, {1}", topicId, parameters);

            var type = parameters.GetValue(0).ToString();
            var defaultValue = DefaultValue;
            string requestId;

            lock (_mappingsLocker)
            {
                switch (type)
                {
                    case RtdTypes.ConnectionState:
                        requestId = RtdTypes.ConnectionState; // All map to unique
                        break;
                    case RtdTypes.Request:
                        requestId = parameters.GetValue(1).ToString();
                        _requestIdToPayload.TryGetValue(requestId, out string payload);
                        SecretariumFunctions.Scp.Send(payload);
                        break;
                    default:
                        return "Invalid parameters";
                }

                _requesIdToTopics[requestId] = topicId;
                _topicsToRequestId[topicId] = requestId;
            }

            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.DebugFormat(@"RequestRtdServer topic connected {0}, {1}, {2}", topicId, parameters, requestId);

            if (!_timerStarted)
            {
                _timerStarted = true;
                _timer.Start();
            }

            return defaultValue;
        }

        public Array RefreshData(ref int topicCount)
        {
            lock (_updatesLocker)
            {
                CleanUpUpdates();

                topicCount = _updates.Count;
                object[,] data = new object[2, _updates.Count];
                int i = 0;

                foreach (var kvp in _updates)
                {
                    data[0, i] = _requesIdToTopics[kvp.Key];
                    data[1, i] = kvp.Value;
                    ++i;
                }

                _updates.Clear();

                return data;
            }
        }

        public void DisconnectData(int topicId)
        {
            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.DebugFormat(@"RequestRtdServer topic disconnecting {0}", topicId);

            lock (_mappingsLocker)
            {
                if (_topicsToRequestId.ContainsKey(topicId))
                {
                    _requesIdToTopics.Remove(_topicsToRequestId[topicId]);
                    _topicsToRequestId.Remove(topicId);
                }

                if (_topicsToRequestId.Count == 0)
                {
                    _timer.Stop();
                    _timerStarted = false;
                }
            }

            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.DebugFormat(@"RequestRtdServer topic disconnected {0}", topicId);
        }

        public int Heartbeat()
        {
            return 1;
        }

        void IRtdServer.ServerTerminate()
        {
            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.Debug(@"RequestRtdServer terminating");

            lock (_updatesLocker)
            {
                _timer.Dispose();
                _updates.Clear();
            }

            lock (_mappingsLocker)
            {
                _requesIdToTopics.Clear();
                _topicsToRequestId.Clear();
                _xlRequestToRequestId.Clear();
                _requestIdToPayload.Clear();

                _requesIdToTopics = null;
                _topicsToRequestId = null;
                _xlRequestToRequestId = null;
                _requestIdToPayload = null;
            }

            if (SecretariumFunctions.Logger.IsDebugEnabled)
                SecretariumFunctions.Logger.Debug(@"RequestRtdServer terminated");
        }

        private void CleanUpUpdates()
        {
            lock (_updatesLocker)
            {
                var toRemove = _updates.Where(x => !_requesIdToTopics.ContainsKey(x.Key)).Select(x => x.Key).ToList();
                foreach (var key in toRemove)
                {
                    _updates.Remove(key);
                }
            }
        }

        #endregion

        #region Secretarium Notification wrapper

        private static void SetValue(string requestId, object value)
        {
            lock (_updatesLocker)
            {
                _updates[requestId] = value;
            }
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
                    SetValue(requestId, error);
                    return;
                }

                var state = json["state"]?.Value<string>();
                if (!string.IsNullOrEmpty(state))
                {
                    SetValue(requestId, state + " @" + DateTime.Now);
                    return;
                }

                var result = json["result"];
                if (result != null)
                {
                    SetValue(requestId, result.ToString());
                    return;
                }
            }
            catch (Exception ex)
            {
                SecretariumFunctions.Logger.Error("Can't parse", ex);
            }
        }

        /// <summary>
        /// Everytime the value of the calling cell is updated by the Rtd server, Excel calls the UDF again.
        /// We need a way to identify a Request from its parameters to avoid querying the platform again for technical reasons.
        /// </summary>
        internal class XlRequest : IEquatable<XlRequest>
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
                    for (int i = 0; i < Args.Length; i++)
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

        #endregion
    }
}
