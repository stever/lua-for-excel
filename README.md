# Lua for Excel

This project shows how Excel functions using the Lua programming language
can be used in Excel. It is currently provided as a proof-of-concept, and
light on features. It currently supports a single Lua script embedded in
the Excel file as an XML Part. This script is parsed for Lua functions,
which are then registered as functions using Excel-DNA.

There is no API provided currently to the Lua scripts, but there is
potential there. The Lua sandbox is set currently restricted to modules
defined by the MoonSharp Preset_HardSandbox setting.

See // http://www.moonsharp.org/sandbox.html for more info.