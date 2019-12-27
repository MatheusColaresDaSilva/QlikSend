Function convertFromUTF8(sIn)

    Dim oIn: Set oIn = CreateObject("ADODB.Stream")
    oIn.Open
    oIn.CharSet = "WIndows-1252"
    oIn.WriteText sIn
    oIn.Position = 0
    oIn.CharSet = "UTF-8"
    ConvertFromUTF8 = oIn.ReadText
    oIn.Close

End Function