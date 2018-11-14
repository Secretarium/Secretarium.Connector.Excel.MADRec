using ExcelDna.Integration.Rtd;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Secretarium.Excel
{
    public class SecretariumRtdServer : IRtdServer
    {
        private const string DefaultValue = "loading...";

        private bool _timerStarted = false;
        private IRTDUpdateEvent _callback;
        private System.Windows.Forms.Timer _timer;
        private Dictionary<string, int> _requesIdToTopics;
        private Dictionary<int, string> _topicsToRequestId;

        private static object _locker = new object();
        private static Dictionary<string, object> _updates = new Dictionary<string, object>();


        private void CleanUpUpdates()
        {
            lock (_locker)
            {
                var toRemove = _updates.Where(x => !_requesIdToTopics.ContainsKey(x.Key)).Select(x => x.Key).ToList();
                foreach (var key in toRemove)
                {
                    _updates.Remove(key);
                }
            }
        }

        protected ILog Logger
        {
            get { return SecretariumFunctions.Logger; }
        }
        
        public virtual int ServerStart(IRTDUpdateEvent callback)
        {
            if (Logger.IsDebugEnabled)
                Logger.Debug(@"RequestRtdServer starting");

            _callback = callback;
            _requesIdToTopics = new Dictionary<string, int>();
            _topicsToRequestId = new Dictionary<int, string>();

            // This Windows.Forms Single threaded (UI thread) timer must be started in this correct COM appartment
            _timer = new System.Windows.Forms.Timer { Interval = 1000 };
            _timer.Tick += (object sender, EventArgs args) =>
            {
                _timer.Stop();

                lock (_locker)
                {
                    CleanUpUpdates();

                    if (_updates.Count > 0)
                        _callback.UpdateNotify();
                }

                _timer.Start();
            };

            if (Logger.IsDebugEnabled)
                Logger.Debug(@"RequestRtdServer started");

            return 1; // 1 means OK
        }

        void IRtdServer.ServerTerminate()
        {
            if (Logger.IsDebugEnabled)
                Logger.Debug(@"RequestRtdServer terminating");

            _timer.Dispose();
            _requesIdToTopics.Clear();
            _topicsToRequestId.Clear();
            _updates.Clear();

            _timer = null;
            _requesIdToTopics = null;
            _topicsToRequestId = null;

            if (Logger.IsDebugEnabled)
                Logger.Debug(@"RequestRtdServer terminated");
        }

        public object ConnectData(int topicId, ref Array parameters, ref bool newValues)
        {
            if (Logger.IsDebugEnabled)
                Logger.DebugFormat(@"RequestRtdServer topic connecting {0}, {1}", topicId, parameters);

            if (parameters.Length == 0)
                return "Invalid parameters";

            var type = parameters.GetValue(0).ToString();
            var defaultValue = DefaultValue;
            string requestId;
            switch (type)
            {
                case SecretariumConnectionStateRtdServer.Name:
                    requestId = SecretariumConnectionStateRtdServer.Name; // All map to unique
                    break;
                case SecretariumRequestRtdServer.Name:
                    defaultValue = SecretariumRequestRtdServer.DefaultValue;
                    if (!SecretariumRequestRtdServer.OnConnectData(ref parameters, out requestId, out string error))
                        return error;
                    break;
                default:
                    return "Invalid parameters";
            }

            _requesIdToTopics[requestId] = topicId;
            _topicsToRequestId[topicId] = requestId;

            if (Logger.IsDebugEnabled)
                Logger.DebugFormat(@"RequestRtdServer topic connected {0}, {1}, {2}", topicId, parameters, requestId);

            if (!_timerStarted)
            {
                _timerStarted = true;
                _timer.Start();
            }

            return defaultValue;
        }

        public void DisconnectData(int topicId)
        {
            if (Logger.IsDebugEnabled)
                Logger.DebugFormat(@"RequestRtdServer topic disconnecting {0}", topicId);

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

            if (Logger.IsDebugEnabled)
                Logger.DebugFormat(@"RequestRtdServer topic disconnected {0}", topicId);
        }

        public Array RefreshData(ref int topicCount)
        {
            lock (_locker)
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

        public int Heartbeat()
        {
            return 1;
        }

        public static void SetValue(string requestId, object value)
        {
            lock(_locker)
            {
                _updates[requestId] = value;
            }
        }
    }
}
