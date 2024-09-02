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
   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim CompressedData() As Byte = {}
         Dim DecompressedData As New List(Of Byte)
         Dim InputFile As String = Nothing
         Dim OutputFile As String = Nothing

         If GetCommandLineArgs().Count = 2 Then
            InputFile = GetCommandLineArgs().Last()
            OutputFile = $"{InputFile}.dat"
            CompressedData = File.ReadAllBytes(InputFile)
            DecompressedData = DecompressData(CompressedData)

            Console.WriteLine($"Writing to: {OutputFile}")
            If DecompressedData.Count = 0 Then Throw New Exception($"Could not decompress {InputFile}.")
            File.WriteAllBytes(OutputFile, DecompressedData.ToArray())
         Else
            Throw New Exception("No input file specified.")
         End If
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure decompresses the specified compressed data and returns the result.
   Private Function DecompressData(CompressedData() As Byte) As List(Of Byte)
      Try
         Dim ByteO As New Byte
         Dim DecompressedData As New List(Of Byte)
         Dim DecompressedSize As New Integer
         Dim Dictionary(&H0% To &HFFF%) As Byte
         Dim DictionaryPosition As Integer = &HFEE%
         Dim Flag As New Integer
         Dim Flags As New Integer
         Dim Length As New Integer
         Dim Offset As New Integer

         Using CompressedDataSream As New BinaryReader(New MemoryStream(CompressedData))
            DecompressedSize = CompressedDataSream.ReadInt32()
            While (CompressedDataSream.BaseStream.Position < CompressedDataSream.BaseStream.Length)
               ByteO = CompressedDataSream.ReadByte()
               If Flags <= &H1% Then
                  Flags = &H100% Or ByteO
               Else
                  Flag = Flags And &H1%
                  Flags = Flags >> &H1%
                  If Flag = &H0% Then
                     Offset = ByteO
                     ByteO = CompressedDataSream.ReadByte()
                     Offset = Offset Or (ByteO And &HF0%) << &H4%
                     Length = (ByteO And &HF%) + &H3%
                     While Length > &H0%
                        Length -= &H1%
                        ByteO = Dictionary(Offset)
                        Offset = (Offset + &H1%) And &HFFF%
                        DecompressedData.Add(ByteO)
                        Dictionary(DictionaryPosition) = ByteO
                        DictionaryPosition = (DictionaryPosition + &H1%) And &HFFF%
                     End While
                  ElseIf Flag = &H1% Then
                     Dictionary(DictionaryPosition) = ByteO
                     DictionaryPosition = (DictionaryPosition + &H1%) And &HFFF%
                     DecompressedData.Add(ByteO)
                  End If
               End If
            End While
         End Using

         Return If(DecompressedData.Count = DecompressedSize, DecompressedData, New List(Of Byte))
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return New List(Of Byte)
   End Function

   'This procedure displays any errors that occur.
   Private Sub DisplayError(ExceptionO As Exception)
      Try
         Console.Error.WriteLine($"ERROR: {ExceptionO.Message}")
         [Exit](0)
      Catch
         [Exit](0)
      End Try
   End Sub
End Module
