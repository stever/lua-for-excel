using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using ExcelDna.Integration;

namespace LuaForExcel
{
    public partial class LuaEditor : Form
    {
        private const string MainScriptName = "Main";

        private static readonly NetOffice.ExcelApi.Application Excel
            = new NetOffice.ExcelApi.Application(null, ExcelDnaUtil.Application);

        private bool _loading = true;

        public LuaEditor()
        {
            InitializeComponent();
        }

        private void LuaEditor_Load(object sender, EventArgs e)
        {
            try
            {
                fastColoredTextBox1.Text = GetLuaScript(MainScriptName) ?? "";
                _loading = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void fastColoredTextBox1_TextChanged(object sender, FastColoredTextBoxNS.TextChangedEventArgs e)
        {
            if (_loading) return;

            try
            {
                SaveLuaScript(MainScriptName, fastColoredTextBox1.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GetLuaScript(string name)
        {
            var scripts = GetLuaScripts();
            return scripts.ContainsKey(name)
                ? scripts[name]
                : null;
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

        private static void SaveLuaScript(string name, string luaScript)
        {
            using (var workbook = Excel.ActiveWorkbook)
            {
                // Delete existing XML part.
                foreach (var part in workbook.CustomXMLParts)
                {
                    var doc = new XmlDocument();
                    doc.Load(new StringReader(part.XML));
                    var root = doc.DocumentElement;
                    Debug.Assert(root != null);

                    if ((root.Attributes["name"]?.InnerText ?? "Unnamed") != name)
                    {
                        continue;
                    }

                    part.Delete();
                    break;
                }

                // Add as new XML part.
                var newDoc = new XmlDocument();
                var newRoot = newDoc.CreateElement("LuaScript");
                newRoot.SetAttribute("name", name);
                newRoot.AppendChild(newDoc.CreateCDataSection(luaScript));
                newDoc.AppendChild(newRoot);
                workbook.CustomXMLParts.Add(newDoc.OuterXml);
            }
        }
    }
}
