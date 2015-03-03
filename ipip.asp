<% 
    Function B2IU(b)
        B2IU = b(3) + b(2) * 256 + b(1) * 65536 + b(0) * 16777216
    End Function

    Function B2IUS(b, s)
        B2IUS = b(s + 3) + b(s + 2) * 256 + b(s + 1) * 65536 + b(s) * 16777216
    End Function

    Function B2IL(b)
        B2IL = b(0) + b(1) * 256 + b(2) * 65536 + b(3) * 16777216
    End Function

    Function B2ILS(b, s)
        B2ILS = b(s) + b(s + 1) * 256 + b(s + 2) * 65536 + b(s + 3) * 16777216
    End Function
        Function GetDateLastModified(FileName)
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set FILE = FSO.GetFile(FileName)
        GetDateLastModified = FILE.DateLastModified
        Set FILE = Nothing
        Set FSO = Nothing
    End Function

    Function GetDateLastModified(FileName)
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FileExists(FileName) Then
            Set FILE = FSO.GetFile(FileName)
            GetDateLastModified = FILE.DateLastModified
            Set FILE = Nothing
        Else
            GetDateLastModified = 0
        End If
        Set FSO = Nothing
    End Function
    
    Function TransBinary(Buf)
        Dim I, aBuf, Size, bStream
        Size = UBound(Buf): ReDim aBuf(Size \ 2)
        For I = 0 To Size - 1 Step 2
            aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))
        Next
        If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))
        aBuf=Join(aBuf, "")
        Set bStream = CreateObject("ADODB.Stream")
        bStream.Type = 1: bStream.Open
        With CreateObject("ADODB.Stream")
            .Type = 2 : .Open: .WriteText aBuf
            .Position = 2: .CopyTo bStream: .Close
        End With
        bStream.Position = 0
        bStream.Type = 2
        bStream.Charset = "UTF-8"
        Txt = bStream.ReadText
        bStream.Close
        TransBinary = Txt
        Set bStream = Nothing
    End Function

    Function Query(ip)
        DIM ipss, ips, ip_prefix_value, ip2long_value, start, max_comp_len, index_offset, index_length
        ipss = Split(ip, ".")
    
        If uBound(ipss) + 1 <> 4 Then
            Query = "错误的IP地址"
            Exit Function
        End If

        If Not (IsNumeric(ipss(0)) And IsNumeric(ipss(1)) And IsNumeric(ipss(2)) And IsNumeric(ipss(3))) Then
            Query = "错误的IP地址"
            Exit Function
        End If

        ips = Array(CInt(ipss(0)),CInt(ipss(1)),CInt(ipss(2)),CInt(ipss(3)))

        If ips(0) > 255 Or ips(1) > 255 Or ips(2) > 255 Or ips(3) > 255 Then
            Query = "错误的IP地址"
            Exit Function
        End If

        Const IPIP_TIME = "ipip_time"
        Const IPIP_LENGTH = "ipip_length"
        Const IPIP_DATA = "ipip_data"
        Const IPIP_INDEXBUFFER = "ipip_indexbuffer"
        Const IPIP_INDEX = "ipip_index"
    
        Dim FileName
        FileName = Server.MapPath("17monipdb.dat")

        Dim LastModified
        LastModified = GetDateLastModified(FileName)
        
        If LastModified = 0 Then
            Query = "17monipdb.dat数据文件不存在"
            Exit Function
        End If
        
        Dim I
        Dim dataBuffer, indexBuffer, index, indexLength
        If DateDiff("s",CDate(Application(IPIP_TIME)), LastModified) <> 0 Then
            Application.Lock
            With CreateObject("ADODB.Stream")
                .Mode = 3: .Type = 1: .Open: .LoadFromFile FileName
                ReDim dataBuffer(.Size - 1)
                For I = 0 To .Size - 1: dataBuffer(I) = AscB(.Read(1)): Next
                .Close
            End With
            indexLength = B2IU(dataBuffer)

            ReDim indexBuffer(indexLength)
            For I = 0 To indexLength: indexBuffer(I) = dataBuffer(i + 4): Next
    
            ReDim index(256)
            For I = 0 To 255: index(I) = B2ILS(indexBuffer, I * 4): Next
            Application(IPIP_TIME) = LastModified
            Application(IPIP_LENGTH) = indexLength
            Application(IPIP_DATA) = dataBuffer
            Application(IPIP_INDEXBUFFER) = indexBuffer
            Application(IPIP_INDEX) = index
            Application.UnLock
        Else
            indexLength = Application(IPIP_LENGTH)
            index = Application(IPIP_INDEX)
            indexBuffer = Application(IPIP_INDEXBUFFER)
            dataBuffer = Application(IPIP_DATA) 
        End If

        ip_prefix_value = ips(0)
        ip2long_value = B2IU(ips)
        start = index(ip_prefix_value)

        max_comp_len = indexLength - 1028

        index_offset = 0
        index_length = 0

        For I = start * 8 + 1024 To max_comp_len - 1 Step 8
            If B2IUS(indexBuffer, I) >= ip2long_value Then
                index_offset = indexBuffer(I + 4) + indexBuffer(I + 5) * 256 + indexBuffer(I + 6) * 65536 + 0 * 16777216
                index_length = indexBuffer(I + 7)
                Exit For
            End If
        Next

        ReDim areaBytes(index_length)

        For I = 0 To index_length - 1: areaBytes(I) = CByte(dataBuffer(I + indexLength + index_offset - 1024)): Next
        Query = TransBinary(areaBytes)
    End Function

    '例：Query("8.8.8.8")
    Response.Write Query(Request.QueryString("ip"))
%> 