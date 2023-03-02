# OpenXML memory usage hack
A hacky solution for OpenXML using a ton of memory during Excel file export.
<p align="center"><img src="/docs/compare_results.png" width="600"/></p>

# Disclaimer
This is not a proper fix. If you have the luxury of time, wait for the OpenXML people to handle the problem proper. If on the other hand, you can see the deadline in your rearview mirror, well...

# How does it work?
The idea is simple: replace the fat OpenXML `Cell` object with just the value you'll put in the cell. Then, on the fly, create the `Cell` OpenXML object, let the library use it, then drop it again. No reference, GC does the reset.

## How does that work?
Seems like some heavy injection. Due to a lucky turn of events, no. We inherit from the `Row` OpenXML class. There we store all the cell values needed to create the `Cell` on the fly. By overriding just a few `Children`, `HasChildren` etc. properties, we get a near 1:1 replacement for the `Row` class. All with ~50 lines of code.

## In code
The code is based on the standard save example from OpenXML ([link](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-2.8.1 "link")) modified to save a table of strings. From there, it is just replacing `Row` with `RowProxy` and feeding it values for cells instead of `Cell` objects.
<p align="center"><img src="/docs/compare_code.png" width="600"/></p>
