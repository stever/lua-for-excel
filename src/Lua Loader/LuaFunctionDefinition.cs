using System.Collections.Generic;

namespace LuaForExcel.LuaLoader
{
    public class LuaFunctionDefinition
    {
        public string Name { get; }
        public List<string> Args { get; }

        public LuaFunctionDefinition(string name, List<string> args)
        {
            Name = name;
            Args = args ?? new List<string>();
        }
    }
}
