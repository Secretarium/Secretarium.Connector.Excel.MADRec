namespace Secretarium.Excel
{
    public class SecretariumConnectionStateRtdServer
    {
        private static bool _isStarted;

        public const string Name = "ConnectionState";
        public const string DefaultValue = "loading...";
        
        public static void EnsureStarted()
        {
            if (_isStarted) return;

            _isStarted = true;
            SecretariumFunctions.Scp.OnStateChange += x =>
            {
                SecretariumRtdServer.SetValue(Name, SecretariumFunctions.Scp.State.ToString());
            };
        }
    }
}