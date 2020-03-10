using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConnectionExample
{
    class Program
    {

        private const string REG_PATH = @"Software\Sage\MMS";
        private const string REGKEY_VALUE = "ClientInstallLocation";
        private const string DEFAULT_ASSEMBLY = "Sage.Common.dll";
        private const string ASSEMBLY_RESOLVER = "Sage.Common.Utilities.AssemblyResolver";
        private const string RESOLVER_METHOD = "GetResolver";

        static void Main(string[] args)
        {
            // Find the 200 client installation folder
            // Sage.Common.dll should be loaded to allow the external application to run without the need to package the required Sage dependencies.
            // Ensure your Sage assemblies are referenced with the copy local properties set to false. 
            AssemblyResolver();

            // Connect to the company through the application object
            Connect();
        }

        private static void Connect()
        {
            // Declare the Application object used to connect.
            Sage.Accounting.Application application = null;

            try
            {
                application = new Sage.Accounting.Application();

                // Use the Connect method (no parameters required)
                application.Connect();

                // IMPORTANT: select the company (database) - a connection will not be made 
                // unless this line is included.
                // In this example an indexer is used to select the company. This
                // is fine if there is only one company. If there are multiple companies a more
                // appropriate solution would be to iterate the collection checking the Name property. 
                application.ActiveCompany = application.Companies[0];

                System.Text.StringBuilder stringBuilder = new System.Text.StringBuilder();

                stringBuilder.Append("Successfully logged on to " + application.ActiveCompany.Name);

                System.Diagnostics.Debug.WriteLine(stringBuilder.ToString());
            }
            // The Connect method may throw the following exceptions:
            // Ex9990Exception - User already logged in.
            // Ex9991Exception - Maximum number of users logged in.
            catch (System.Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(exception.Message);
            }
            finally
            {
                if (application != null)
                {
                    // Disconnect from the application
                    application.Disconnect();
                }
            }
        }

        private static void AssemblyResolver()
        {
            // get registry info for Sage 200 server path
            string path = string.Empty;
            Microsoft.Win32.RegistryKey root = Microsoft.Win32.Registry.CurrentUser;
            Microsoft.Win32.RegistryKey key = root.OpenSubKey(REG_PATH);

            if (key != null)
            {
                object value = key.GetValue(REGKEY_VALUE);
                if (value != null)
                    path = value as string;
            }

            // refer to all installed assemblies based on location of default one
            if (string.IsNullOrEmpty(path) == false)
            {
                string commonDllAssemblyName = System.IO.Path.Combine(path, DEFAULT_ASSEMBLY);

                if (System.IO.File.Exists(commonDllAssemblyName))
                {
                    System.Reflection.Assembly defaultAssembly = System.Reflection.Assembly.LoadFrom(commonDllAssemblyName);
                    Type type = defaultAssembly.GetType(ASSEMBLY_RESOLVER);
                    System.Reflection.MethodInfo method = type.GetMethod(RESOLVER_METHOD);
                    method.Invoke(null, null);
                }
            }

        }


    }

}

