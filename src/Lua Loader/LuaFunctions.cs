using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using Antlr4.Runtime;
using LuaForExcel.LuaLoader.Parser;
using log4net;
using MoonSharp.Interpreter;
using MoonSharp.Interpreter.Loaders;

namespace LuaForExcel.LuaLoader
{
    public class LuaFunctions
    {
        private static readonly ILog Log = LogManager.
            GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private readonly CoreModules _modules;
        private readonly IScriptLoader _scriptLoader;
        private readonly Dictionary<string, Script> _functions =
            new Dictionary<string, Script>(StringComparer.OrdinalIgnoreCase);

        public LuaFunctions(IScriptLoader scriptLoader = null)
        {
            _scriptLoader = scriptLoader;

            // http://www.moonsharp.org/sandbox.html
            _modules = CoreModules.Preset_HardSandbox;
            if (scriptLoader != null)
            {
                _modules |= CoreModules.LoadMethods; // provides 'require'
            }
        }

        public void Load(string luaScript, string scriptName)
        {
            Log.DebugFormat("Loading Lua script: {0}", scriptName ?? "(no name)");
            if (string.IsNullOrEmpty(luaScript)) return;

            var script = new Script(_modules) {Options =
            {
                DebugPrint = s => Log.DebugFormat("{0}: {1}", scriptName ?? "Lua print", s),
                CheckThreadAccess = false,
                ScriptLoader = _scriptLoader
            }};

            script.DoString(luaScript);

            foreach (var def in GetFunctionDefinitions(luaScript))
            {
                if (_functions.ContainsKey(def.Name))
                {
                    Log.Warn($"Ignoring redefined Lua function: {def.Name}");
                    continue;
                }

                Log.DebugFormat("Registering Lua function: {0}", def.Name);
                _functions.Add(def.Name, script);
            }
        }

        public DynValue RunFunction(string name, params object[] args)
        {
            Log.DebugFormat("Lua function call: {0}", GetFunctionCall(name, args));
            try
            {
                var script = _functions[name];
                return script.Call(script.Globals[name], args);
            }
            catch (Exception ex)
            {
                // Detail the function call in the exception message.
                throw new Exception($"Lua function call failed: {GetFunctionCall(name, args)}", ex);
            }
        }

        public static IEnumerable<LuaFunctionDefinition> GetFunctionDefinitions(string luaScript)
        {
            var inputStream = new AntlrInputStream(luaScript);
            var lexer = new LuaLexer(inputStream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new LuaParser(tokens);

            var visitor = new LuaFunctionsVisitor();
            visitor.VisitChunk(parser.chunk());
            return visitor.FunctionDefinitions;
        }

        private static string GetFunctionCall(string name, params object[] args)
        {
            var sb = new StringBuilder();
            sb.Append($"{name}(");
            if (args != null)
            {
                for (var i = 0; i < args.Length; i++)
                {
                    var arg = args[i];
                    switch (arg)
                    {
                        case int _:
                        case double _:
                            sb.Append(arg);
                            break;
                        case string str:
                            sb.Append($"\"{str}\"");
                            break;
                        case bool b:
                            sb.Append(b ? "true" : "false");
                            break;
                        case null:
                            sb.Append("null");
                            break;
                        default:
                            Log.WarnFormat("Unexpected arg type in function call: {0}", arg.GetType().Name);
                            sb.Append(arg);
                            break;
                    }

                    if (i < args.Length - 1)
                    {
                        sb.Append(", ");
                    }
                }
            }
            sb.Append(")");
            return sb.ToString();
        }
    }
}
