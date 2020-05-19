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
Using `MSXML2.ServerXMLHTTP60`

Partly inspired by:
  * https://github.com/andreafortuna-org/VBAIPFunctions


== Getting Started
Make sure to enable `Microsoft XML 60` in Tools > References.

