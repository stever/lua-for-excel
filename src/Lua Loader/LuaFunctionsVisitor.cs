using System.Collections.Generic;
using System.Linq;
using LuaForExcel.LuaLoader.Parser;

namespace LuaForExcel.LuaLoader
{
    internal class LuaFunctionsVisitor : LuaBaseVisitor<object>
    {
        public List<LuaFunctionDefinition> FunctionDefinitions { get; } = new List<LuaFunctionDefinition>();

        public override object VisitStatFunction(LuaParser.StatFunctionContext context)
        {
            var name = context.functionName.GetText();
            var args = context.functionBody.parameterList?.parameterNames?.GetText().Split(',').ToList();
            FunctionDefinitions.Add(new LuaFunctionDefinition(name, args));
            return base.VisitStatFunction(context);
        }
    }
}
