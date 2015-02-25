Public Class MultiIcon

        Public Count As Integer
        Public Icons As New List(Of Icon)
        Private Size As Integer
        Private MultiIcons() As String
        Private EntryIcons As New List(Of IconEntry)
        Private Stream As MemoryStream

        Structure IconEntry
            Public Width As Byte
            Public Height As Byte
            Public ColorCount As Byte
            Public Reserved As Byte
            Public Planes As Short
            Public BitCount As Short
            Public BytesInRes As Integer
            Public ImageOffset As Integer
        End Structure

        Sub New(Icons() As String)
            MultiIcons = Icons
            Count = Icons.Count
        End Sub

        Sub New(File As String)
            Dim Bytes = IO.File.ReadAllBytes(File)

            Stream = New MemoryStream(Bytes)
            Stream.Seek(6, SeekOrigin.Begin)

            Count = BitConverter.ToInt16(Bytes, 4)

            Dim Ie As IconEntry
            For i = 0 To Count - 1
                Dim Br As New BinaryReader(Stream)
                Ie.Width = Br.ReadByte
                Ie.Height = Br.ReadByte
                Ie.ColorCount = Br.ReadByte
                Ie.Reserved = Br.ReadByte
                Ie.Planes = Br.ReadInt16
                Ie.BitCount = Br.ReadInt16
                Ie.BytesInRes = Br.ReadInt32
                Ie.ImageOffset = Br.ReadInt32
                EntryIcons.Add(Ie)
            Next

            EntryIcons = EntryIcons.OrderBy(Function(i) i.Width).ToList

            For i = 0 To Count - 1
                Icons.Add(GetIcon(i))
            Next
        End Sub

        Public Sub SaveIcon(Index As Integer, File As String)
            Dim Fs As New FileStream(File, FileMode.OpenOrCreate, FileAccess.Write)
            Icons(Index).Save(Fs)
            Fs.Close()
        End Sub

        Public Sub SaveIcons(Path As String)
            For i = 0 To Count - 1
                Dim Fs As New FileStream(Path & "\" & i & ".ico", FileMode.OpenOrCreate, FileAccess.Write)
                Icons(i).Save(Fs)
                Fs.Close()
            Next
        End Sub

        Private Function GetIcon(Index As Integer) As Icon
            Dim Ie As IconEntry = EntryIcons(Index)

            Dim Ms As New MemoryStream
            Dim Bw As New BinaryWriter(Ms)
            Bw.Write(CShort(0))
            Bw.Write(CShort(1))
            Bw.Write(CShort(1))
            Bw.Write(Ie.Width)
            Bw.Write(Ie.Height)
            Bw.Write(Ie.ColorCount)
            Bw.Write(Ie.Reserved)
            Bw.Write(Ie.Planes)
            Bw.Write(Ie.BitCount)
            Bw.Write(Ie.BytesInRes)
            Bw.Write(22)

            Dim Bytes(Ie.BytesInRes - 1) As Byte
            Stream.Seek(Ie.ImageOffset, SeekOrigin.Begin)
            Stream.Read(Bytes, 0, Ie.BytesInRes)
            Bw.Write(Bytes)

            Ms.Seek(0, SeekOrigin.Begin)
            Return New Icon(Ms, Ie.Width, Ie.Height)
            Bw.Close()
            Stream.Close()
        End Function

        Public Sub SaveMultiIcon(File As String)
            Dim Fs As New FileStream(File, FileMode.OpenOrCreate, FileAccess.Write)
            MultiIcon.Save(Fs)
            Fs.Close()
        End Sub

        Private Function MultiIcon() As Icon
            Dim Ms As New MemoryStream
            Dim Bw As New BinaryWriter(Ms)

            Bw.Write(CShort(0))
            Bw.Write(CShort(1))
            Bw.Write(CShort(Count))

            Dim ImageOffset As Integer = 6 + (16 * Count)
            Dim ByteIcons As New List(Of Byte())

            Dim Ie As IconEntry
            For i = 0 To Count - 1
                Dim Bytes = IO.File.ReadAllBytes(MultiIcons(i))

                Stream = New MemoryStream(Bytes)
                Stream.Seek(6, SeekOrigin.Begin)

                Dim Br As New BinaryReader(Stream)
                Ie.Width = Br.ReadByte
                If i = 0 Then Size = CInt(Ie.Width)
                Ie.Height = Br.ReadByte
                Ie.ColorCount = Br.ReadByte
                Ie.Reserved = Br.ReadByte
                Ie.Planes = Br.ReadInt16
                Ie.BitCount = Br.ReadInt16
                Ie.BytesInRes = Br.ReadInt32
                Ie.ImageOffset = Br.ReadInt32

                Bw.Write(CByte(0))
                Bw.Write(CByte(0))
                Bw.Write(Ie.ColorCount)
                Bw.Write(Ie.Reserved)
                Bw.Write(Ie.Planes)
                Bw.Write(Ie.BitCount)
                Bw.Write(Ie.BytesInRes)
                Bw.Write(ImageOffset)
                ImageOffset += Ie.BytesInRes

                Dim Byt(Ie.BytesInRes - 1) As Byte
                Stream.Seek(Ie.ImageOffset, SeekOrigin.Begin)
                Stream.Read(Byt, 0, Ie.BytesInRes)

                ByteIcons.Add(Byt)
            Next

            For Each b In ByteIcons
                Bw.Write(b)
            Next

            Ms.Seek(0, SeekOrigin.Begin)
            Return New Icon(Ms, Size, Size)
            Bw.Close()
            Stream.Close()
        End Function
    End Class
