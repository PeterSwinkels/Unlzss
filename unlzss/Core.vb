'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module CoreModule
   Private Const MAXIMUM_LENGTH As Integer = &H12%   'Defines the maximum size for a chunk of decompressed data to be copied from the buffer.
   Private Const MINIMUM_LENGTH As Integer = &H2%    'Defines the minimum size for a chunk of decompressed data to be copied from the buffer.
   Private Const RING_SIZE As Integer = &H1000%      'Defines the last chunk's offset inside the buffer.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim CompressedData As New List(Of Byte)
         Dim DecompressedData As New List(Of Byte)
         Dim DecompressedSize As New Integer
         Dim InputFile As String = Nothing
         Dim OutputFile As String = Nothing

         If GetCommandLineArgs().Count = 2 Then
            InputFile = GetCommandLineArgs().Last()
            OutputFile = $"{InputFile}.dat"
            CompressedData = New List(Of Byte)(File.ReadAllBytes(InputFile))
            DecompressedSize = BitConverter.ToInt32(CompressedData.GetRange(0, 4).ToArray(), 0)
            DecompressedData = DecompressLZSS(CompressedData.GetRange(4, CompressedData.Count - 4), DecompressedSize)

            Console.WriteLine($"Writing to: {OutputFile}")
            If Not DecompressedData.Count = DecompressedSize Then Throw New Exception($"Could not decompress {InputFile}.")
            File.WriteAllBytes(OutputFile, DecompressedData.ToArray())
         Else
            Throw New Exception("No input file specified.")
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure decompresses the specified LZSS compressed data and returns the decompressed data if successful.
   Private Function DecompressLZSS(CompressedData As List(Of Byte), DecompressedSize As Integer) As List(Of Byte)
      Try
         Dim Buffer(&H0 To (RING_SIZE + MAXIMUM_LENGTH) - &H2%) As Byte
         Dim BufferIndex As Integer = RING_SIZE - MAXIMUM_LENGTH
         Dim ByteO As New Byte
         Dim Count As New Integer
         Dim DecompressedData As New List(Of Byte)
         Dim Flags As Integer = &H0%
         Dim Offset As New Integer
         Dim Position As Integer = &H0%

         Do
            Flags = Flags >> &H1%
            If (Flags And &H100%) = &H0% Then
               If Not GetByte(CompressedData, ByteO, Position) Then Exit Do
               Flags = ByteO Or &HFF00%
            End If
            If (Flags And &H1%) = &H1% Then
               If Not GetByte(CompressedData, ByteO, Position) Then Exit Do

               If (DecompressedData.Count >= DecompressedSize) Then Exit Do
               DecompressedData.Add(ByteO)

               Buffer(BufferIndex) = ByteO
               BufferIndex = (BufferIndex + &H1%) And (RING_SIZE - &H1%)
            Else
               If Not GetByte(CompressedData, ByteO, Position) Then Exit Do
               Offset = ByteO

               If Not GetByte(CompressedData, ByteO, Position) Then Exit Do
               Count = ByteO

               Offset = Offset Or ((Count And &HF0%) << &H4%)
               Count = (Count And &HF%) + MINIMUM_LENGTH

               For Index As Integer = &H0% To Count
                  ByteO = (Buffer((Offset + Index) And (RING_SIZE - &H1%)))
                  Buffer(BufferIndex) = ByteO
                  BufferIndex = (BufferIndex + &H1%) And (RING_SIZE - &H1%)

                  If DecompressedData.Count >= DecompressedSize Then Exit Do
                  DecompressedData.Add(ByteO)
               Next Index
            End If
         Loop

         Return If(Position >= CompressedData.Count, DecompressedData, New List(Of Byte))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of Byte)
   End Function

   'This procedure returns the byte at the specified offset from the specified bytes and increments the offset.
   Private Function GetByte(Bytes As List(Of Byte), ByRef ByteO As byte, ByRef Position As Integer) As Boolean
      Try
         If Position < Bytes.Count Then
            ByteO = Bytes(Position)
            Position += 1
            Return True
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure handles any errors that occur.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.WriteLine($"ERROR: {ExceptionO.Message}")
         [Exit](0)
      Catch
         [Exit](0)
      End Try
   End Sub
End Module
