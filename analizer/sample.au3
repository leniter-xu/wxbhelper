#include "analizer.au3"

opt("WinTitleMatchMode", 2)
winwait("SciTE", "")
winactivate("SciTE", "")
AZPrintWindow()
AZTextClick("Search")
AZPrintWindow()
AZTextClick("Find...")
winwait("Find", "")
winactivate("Find", "")
send("Search")
AZPrintWindow()
AZTextClick("Find Next")