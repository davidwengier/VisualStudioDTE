using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace VisualStudioDTE
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (var dte in AutomateVS.FindDTEs())
            {
                // Make sure we can talk to it.. this is just here to throw an exception early
                dte.StatusBar.Text = "Hello World!";

                foreach (var proj in dte.Solution.Projects)
                {
                    try
                    {

                        EnvDTE.Project project = proj as EnvDTE.Project;
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine(project.Name + ":");

                        //Console.WriteLine("Automatic:");
                        Console.ResetColor();
                        EnvDTE.Properties properties = project.Properties;

                        //foreach (var prop in properties)
                        //{
                        //    WriteProperty(prop);
                        //}

                        //Console.ForegroundColor = ConsoleColor.White;
                        //Console.WriteLine("Manually:");
                        //Console.ResetColor();

                        // Make sure we can get some by name
                        //WriteProperty(properties, "OutputType");
                        //WriteProperty(properties, "OutputTypeEx");
                        //WriteProperty(properties, "TargetFramework");
                        //WriteProperty(properties, "TargetFrameworkMoniker");
                        //WriteProperty(properties, "TargetFrameworkMonikers");
                        //WriteProperty(properties, "Authors");
                        //WriteProperty(properties, "PreBuildEvent");
                        //WriteProperty(properties, "PostBuildEvent");
                        //WriteProperty(properties, "ApplicationManifest");
                        //WriteProperty(properties, "PackageId");
                        //WriteProperty(properties, "Authors");
                        //WriteProperty(properties, "AssemblyOriginatorKeyFile");
                        //WriteProperty(properties, "RunCodeAnalysis");

                        var serviceProvider = dte.Application as Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

                        var exprEvaluator = (IVsBooleanSymbolExpressionEvaluator)GetService(serviceProvider, Guid.Parse("59252755-82AC-4A88-A489-453FEEBC694D"), Guid.Parse("7C8306FC-AFBF-43FA-88FC-6FE4D7E16D74"));
                        Console.WriteLine(exprEvaluator.EvaluateExpression("VB + (WPF | (!CPS + WinForms))", "VB CPS WindowsForms"));

                        //var solutionService = (IVsSolution)GetService(serviceProvider, typeof(SVsSolution), typeof(IVsSolution));
                        //var hier = GetHierarchy(solutionService, proj as Project);

                        //WriteHierarchyProperty<__VSDESIGNER_HIDDENCODEGENERATION>(hier, __VSHPROPID2.VSHPROPID_DesignerHiddenCodeGeneration);
                        //WriteHierarchyProperty<__VSPROJOUTPUTTYPE>(hier, __VSHPROPID5.VSHPROPID_OutputType);
                        //WriteHierarchyProperty<VSDESIGNER_FUNCTIONVISIBILITY>(hier, __VSHPROPID.VSHPROPID_DesignerFunctionVisibility);
                        //WriteHierarchyProperty<VSDESIGNER_VARIABLENAMING>(hier, __VSHPROPID.VSHPROPID_DesignerVariableNaming);
                        //WriteHierarchyProperty<VSDESIGNER_VARIABLENAMING>(hier, __VSHPROPID5.VSHPROPID_TargetRuntime);

                        //properties = project.ConfigurationManager.ActiveConfiguration.Properties;

                        //WriteProperty(properties, "RunCodeAnalysis");

                        //foreach (var item in project.ProjectItems.OfType<ProjectItem>())
                        //{
                        //    Console.WriteLine("Properties of " + item.Name);
                        //    WriteProperty(item.Properties, "BuildAction");

                        //    foreach (var prop in item.Properties)
                        //    {
                        //        WriteProperty(prop);
                        //    }
                        //}
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

        private static void WriteHierarchyProperty<TEnum>(IVsHierarchy hier, Enum propId)
        {
            try
            {
                hier.GetProperty((uint)VSConstants.VSITEMID.Root, (int)(object)propId, out object result);
                Console.Write($"{propId.ToString()}:");
                Console.WriteLine((TEnum)result);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: " + ex.Message);
                Console.ResetColor();
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

        private static IVsHierarchy GetHierarchy(IVsSolution solutionService, EnvDTE.Project project)
        {
            solutionService.GetProjectOfUniqueName(project.UniqueName, out IVsHierarchy projectHierarchy);
            return projectHierarchy;
        }

        private static object GetService(Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider, System.Type serviceType, System.Type interfaceType)
        {
            return GetService(serviceProvider, serviceType.GUID, interfaceType.GUID);
        }

        private static object GetService(Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider, Guid serviceGuid, Guid interfaceGuid)
        {
            object service = null;
            IntPtr servicePointer;

            int hr = serviceProvider.QueryService(ref serviceGuid, ref interfaceGuid, out servicePointer);
            if (hr != VSConstants.S_OK)
            {
                System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr);
            }
            else if (servicePointer != IntPtr.Zero)
            {
                service = System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(servicePointer);
                System.Runtime.InteropServices.Marshal.Release(servicePointer);
            }
            return service;
        }
    }
}
