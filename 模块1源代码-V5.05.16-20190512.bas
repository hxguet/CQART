Sub DownFile()
    Dim H, S
    Set H = CreateObject("Microsoft.XMLHTTP")
    H.Open "GET", "https://github.com/hxguet/Course-Quality-Analysis-Report-Template/blob/Course-Quality-Analysis-Report-Template/%E6%A8%A1%E5%9D%971%E6%BA%90%E4%BB%A3%E7%A0%81-V5.05.13-20190509.bas", False
    H.send
    Set S = CreateObject("ADODB.Stream")
    S.Type = 1
    S.Open
    S.Write H.Responsebody
    S.savetofile ("D:/1.bas")
    
End Sub

