VBA = \
r"""

Sub WriteBytes(objFile, strBytes)
    Dim aNumbers
    Dim iIter
	Dim str

    aNumbers = Split(strBytes)
    For iIter = LBound(aNumbers) To UBound(aNumbers)
		s = Hex(aNumbers(iIter))
		If Len(s) < 2 Then 
			s = "0" & s
		End if
        str = str & s
    Next
	Dim stream, xmldom, node
    Set xmldom = CreateObject("Microsoft.XMLDOM")
    Set node = xmldom.CreateElement("binary")
    node.DataType = "bin.hex"
    node.Text = str
	objFile.write node.NodeTypedValue
End Sub

"""