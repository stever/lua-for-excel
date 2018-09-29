using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using ExcelDna.Integration;
using log4net;
using MoonSharp.Interpreter;

namespace LuaForExcel.LuaLoader
{
    public class AddIn : IExcelAddIn
    {
        private static readonly ILog Log = LogManager.
            GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private static readonly NetOffice.ExcelApi.Application Excel
            = new NetOffice.ExcelApi.Application(null, ExcelDnaUtil.Application);

        private static readonly Dictionary<string, Script> Scripts;

        static AddIn()
        {
            Log.Debug("static AddIn");
            Scripts = new Dictionary<string, Script>(StringComparer.OrdinalIgnoreCase);
        }

        public void AutoOpen()
        {
            Log.Debug("AutoOpen");

            Scripts.Clear();

            // Registration takes a list of Delegates
            var delegates = new List<Delegate>();
            var functionAttributes = new List<object>();
            var argAttributesLists = new List<List<object>>();

            try
            {
                var luaScripts = GetLuaScripts();
                foreach (var obj in luaScripts)
                {
                    var scriptName = obj.Key;
                    var luaScript = obj.Value;

                    // http://www.moonsharp.org/sandbox.html
                    var script = new Script(CoreModules.Preset_HardSandbox) {Options =
                    {
                        DebugPrint = s => Log.InfoFormat("{0}: {1}", scriptName ?? "Lua print", s),
                        CheckThreadAccess = false
                    }};

                    script.DoString(luaScript);

                    foreach (var def in LuaFunctions.GetFunctionDefinitions(luaScript))
                    {
                        if (Scripts.ContainsKey(def.Name))
                        {
                            Log.WarnFormat("Ignoring redefined Lua function: {0}", def.Name);
                            continue;
                        }

                        if (def.Args.Count > 16)
                        {
                            Log.ErrorFormat("Functions with more than 16 arguments cannot be registered with Excel-DNA: {0} ({1})", def.Name, def.Args.Count);
                            continue;
                        }

                        // Add the function script to be executed.
                        Log.InfoFormat("Function: {0}", def.Name);
                        Scripts.Add(def.Name, script);

                        // Add a delegate for the function call.
                        delegates.Add(GetDelegate(def.Name, def.Args.Count));

                        // Add the function attribute.
                        // https://github.com/Excel-DNA/ExcelDna/wiki/ExcelFunction-and-other-attributes
                        functionAttributes.Add(new ExcelFunctionAttribute
                        {
                            Name = def.Name,
                            Description = "",
                            IsVolatile = true,
                            IsHidden = false,
                            IsThreadSafe = false,
                            IsClusterSafe = false,
                            IsExceptionSafe = false,
                            IsMacroType = false
                        });

                        // Add the function argument attributes.
                        // Gather list of attributes for this function args.
                        var argAttributes = new List<object>();
                        foreach (var arg in def.Args)
                        {
                            var argAttribute = new ExcelArgumentAttribute
                            {
                                Name = arg,
                                Description = ""
                            };

                            argAttributes.Add(argAttribute);
                        }

                        // Add the function args attributes.
                        argAttributesLists.Add(argAttributes);
                    }
                }

                ExcelIntegration.RegisterDelegates(delegates, functionAttributes, argAttributesLists);
            }
            catch (Exception ex)
            {
                Log.Error("EXCEPTION", ex);
                MessageBox.Show(ex.Message, "AutoOpen Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AutoClose()
        {

        }

        private static Dictionary<string, string> GetLuaScripts()
        {
            var scripts = new Dictionary<string, string>();

            using (var workbook = Excel.ActiveWorkbook)
            {
                foreach (var part in workbook.CustomXMLParts)
                {
                    // Load the XML and check what the root element name is.
                    var doc = new XmlDocument();
                    doc.Load(new StringReader(part.XML));
                    var root = doc.DocumentElement;
                    Debug.Assert(root != null);

                    switch (root.Name)
                    {
                        case "LuaScript":
                            var name = root.Attributes["name"]?.InnerText ?? "Unnamed";
                            var luaScript = root.InnerText;
                            Debug.Assert(!scripts.ContainsKey(name));
                            scripts.Add(name, luaScript);
                            break;
                    }
                }

                return scripts;
            }
        }

        private static object RunFunction(string name, params object[] args)
        {
            try
            {
                Log.DebugFormat("RunFunction {0}", name);
                var script = Scripts[name];
                var result = script.Call(script.Globals[name], args);
                return result.ToObject();
            }
            catch (Exception ex)
            {
                Log.Error("EXCEPTION", ex);
                throw;
            }
        }

        private static Delegate GetDelegate(string name, int argCount)
        {
            switch (argCount)
            {
                case 0: return (Func<object>)(() => RunFunction(name));
                case 1: return (Func<object, object>)(arg => RunFunction(name, arg));
                case 2: return (Func<object, object, object>)((arg1, arg2) => RunFunction(name, arg1, arg2));
                case 3: return (Func<object, object, object, object>)((arg1, arg2, arg3) => RunFunction(name, arg1, arg2, arg3));
                case 4: return (Func<object, object, object, object, object>)((arg1, arg2, arg3, arg4) => RunFunction(name, arg1, arg2, arg3, arg4));
                case 5: return (Func<object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5) => RunFunction(name, arg1, arg2, arg3, arg4, arg5));
                case 6: return (Func<object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6));
                case 7: return (Func<object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7));
                case 8: return (Func<object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8));
                case 9: return (Func<object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9));
                case 10: return (Func<object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10));
                case 11: return (Func<object, object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11));
                case 12: return (Func<object, object, object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12));
                case 13: return (Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13));
                case 14: return (Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14));
                case 15: return (Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15));
                case 16: return (Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object, object, object, object>)((arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16) => RunFunction(name, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16));
                // NOTE: It is not possible to register a function in Excel-DNA with more than 16 arguments.
                default: throw new ArgumentOutOfRangeException($"Unsupported number of args: {argCount}");
            }
        }
    }
}
