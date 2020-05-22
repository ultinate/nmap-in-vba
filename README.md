= Nmap for the Poor

== Why

If you ever find yourself in a situation where you want to scan
a network range and only have VBA scripting available.

== What

Features:

  * Scan an IP range for HTTP(S) responses
  * Slow by design to avoid detection by IDS
    (honestly, I just could make it work faster)

Does not support (yet):

  * Other protocols and ports
  * IPv6
  * CIDR notation
  * Faster scanning

== How

Call to URL is implemented using Microsoft XML  `MSXML2.ServerXMLHTTP60`

Partly inspired by:
  * https://github.com/andreafortuna-org/VBAIPFunctions


== Getting Started

  * In Excel Worksheet, press Alt+F11 to enter VBA Editor. 
  * Inside Excel VBA Editor, go to Tools > References.
  * Enable `Microsoft XML 60`. 
  * In VBA Project Browser, create new "Module". Copy VBA code into this module.
  * In Excel Worksheet, run Macros from Developer Toolbar in Ribbon.
    Make sure you are not in cell editing mode, i.e. you did not double
    click a cell to edit its content.

