using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using System.Reflection;

namespace Todo
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            try
            {
                // Clear Trace listeners to suppress internal EventSource error messages
                Trace.Listeners.Clear();

                // Some platforms may expose Debug.Listeners; clear it if present via reflection to be safe
                var debugType = typeof(Debug);
                var listenersProp = debugType.GetProperty("Listeners", BindingFlags.Static | BindingFlags.Public);
                if (listenersProp != null)
                {
                    var listenersObj = listenersProp.GetValue(null);
                    // Attempt to clear if it has a Clear method
                    var clearMethod = listenersObj?.GetType().GetMethod("Clear", BindingFlags.Instance | BindingFlags.Public);
                    clearMethod?.Invoke(listenersObj, null);
                }
            }
            catch
            {
                // Swallow any failure to avoid impacting app startup
            }
        }
    }
}
