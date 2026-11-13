Option Explicit

Public Sub RunUtilsTests()
    ResetCounters
    Dim g1 As String, g2 As String
    g1 = NewGuid(): g2 = NewGuid()
    AssertTrue Len(g1) = 36, "GUID longueur 36"
    AssertTrue g1 <> g2, "GUID uniques"
    AssertEquals "2026-11-13", FormatIsoDate(DateSerial(2026, 11, 13)), "format ISO"
    AssertTrue ParseIsoDate("2026-11-13") = DateSerial(2026, 11, 13), "parse ISO"
    AssertEquals "a&amp;b&lt;c&gt;d&quot;e&apos;f", EscapeXml("a&b<c>d""e'f"), "échappement XML"
    PrintSummary "Utils"
End Sub