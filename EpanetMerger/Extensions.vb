Imports System.Globalization
Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function ToCamelCase(ByVal aValue As String) As String
        ' Test for nothing or empty.
        If String.IsNullOrEmpty(aValue) Then Return aValue

        Return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(aValue.ToLower())
    End Function
End Module