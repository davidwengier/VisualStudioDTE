using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace VisualStudioDTE
{
    class Program
    {
        static void Main(string[] args)
        {
            var dte = AutomateVS.FindDTE();

            // Make sure we can talk to it.. this is just here to throw an exception early
            dte.StatusBar.Text = "Hello World!";

            foreach (var proj in dte.ActiveSolutionProjects)
            {
                EnvDTE.Project project = proj as EnvDTE.Project;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(project.Name + ":");
                Console.WriteLine("Automatic:");
                Console.ResetColor();
                EnvDTE.Properties properties = project.Properties;
                foreach (var prop in properties)
                {
                    WriteProperty(prop);
                }
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Manually:");
                Console.ResetColor();
                // Make sure we can get some by name
                WriteProperty(properties, "OutputType");
                WriteProperty(properties, "OutputTypeEx");
                WriteProperty(properties, "TargetFramework");
                WriteProperty(properties, "TargetFrameworkMoniker");
                WriteProperty(properties, "TargetFrameworkMonikers");
                WriteProperty(properties, "Authors");
                WriteProperty(properties, "PreBuildEvent");
                WriteProperty(properties, "PostBuildEvent");
                WriteProperty(properties, "ApplicationManifest");
                WriteProperty(properties, "PackageId");
                WriteProperty(properties, "Authors");
                WriteProperty(properties, "AssemblyOriginatorKeyFile");
            }
        }

        private static void WriteProperty(Properties properties, string propName)
        {
            try
            {
                var prop = properties.Item(propName);
                if (prop == null)
                {
                    Console.WriteLine(propName + ": Couldn't find property");
                }
                else
                {
                    WriteProperty(prop);
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(propName + ": ERROR: " + ex.Message);
                Console.ResetColor();
            }
        }

        private static void WriteProperty(object prop)
        {
            EnvDTE.Property property = prop as EnvDTE.Property;
            Console.Write("  " + property.Name + ": ");
            try
            {
                var value = (object)property.Value;
                if (value == null)
                {
                    Console.WriteLine("[null]");
                }
                else
                {
                    Console.Write("(" + value.GetType().Name + ") ");
                    Console.WriteLine(value);
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: " + ex.Message);
                Console.ResetColor();
            }
        }
    }
}
