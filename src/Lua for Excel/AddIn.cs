using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using ExcelDna.Integration;
using log4net;

namespace LuaForExcel
{
    public class AddIn : IExcelAddIn
    {
        private static readonly ILog Log = LogManager.
            GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private static readonly NetOffice.ExcelApi.Application Excel
            = new NetOffice.ExcelApi.Application(null, ExcelDnaUtil.Application);

        private static string _luaLoaderXllFilename;
        private static object _luaLoaderAddInRegistrationId;

        static AddIn()
        {
            Log.Debug("static AddIn");
            Application.EnableVisualStyles();
        }

        public void AutoOpen()
        {
            Log.Debug("AutoOpen");

            try
            {
                var xllPath = Path.GetDirectoryName((string) XlCall.Excel(XlCall.xlGetName));
                Debug.Assert(xllPath != null);
                Log.DebugFormat("XLL path: {0}", xllPath);
                _luaLoaderXllFilename = Path.Combine(xllPath, "Lua Loader-AddIn.xll");
                Log.DebugFormat("Lua loader XLL filename: {0}", _luaLoaderXllFilename);

                Excel.WorkbookOpenEvent += wb => Excel.Run("ReloadLuaLoaderAddIn");
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

        [ExcelFunction(IsHidden = true)]
        public static void ReloadLuaLoaderAddIn()
        {
            try
            {
                if (_luaLoaderAddInRegistrationId != null)
                {
                    Excel.Run("UnloadLuaLoaderAddIn");
                }

                Excel.Run("LoadLuaLoaderAddIn");
            }
            catch (Exception ex)
            {
                Log.Error("EXCEPTION", ex);
                MessageBox.Show(ex.Message, "Reload Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ExcelFunction(IsHidden = true)]
        public static void LoadLuaLoaderAddIn()
        {
            Log.Debug("LoadLuaLoaderAddIn");
            try
            {
                XlCall.TryExcel(XlCall.xlfRegister,
                    out _luaLoaderAddInRegistrationId, _luaLoaderXllFilename);
            }
            catch (Exception ex)
            {
                Log.Error("EXCEPTION", ex);
                MessageBox.Show(ex.Message, "Load Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ExcelFunction(IsHidden = true)]
        public static void UnloadLuaLoaderAddIn()
        {
            Log.Debug("UnloadLuaLoaderAddIn");
            try
            {
                var removeId = XlCall.Excel(XlCall.xlfRegister, _luaLoaderXllFilename,
                    "xlAutoRemove", "I" , ExcelMissing.Value, ExcelMissing.Value, 2);

                XlCall.Excel(XlCall.xlfCall, removeId);
                XlCall.Excel(XlCall.xlfUnregister, removeId);

                /*var success = */XlCall.Excel(XlCall.xlfUnregister,
                    _luaLoaderAddInRegistrationId);
            }
            catch (Exception ex)
            {
                Log.Error("EXCEPTION", ex);
                MessageBox.Show(ex.Message, "Unload Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
