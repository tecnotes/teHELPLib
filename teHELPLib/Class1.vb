Imports System.IO

Public Class INIFile

    Public Function GetValue(ByVal INIFileName As String, ByVal INIOption As String) As String
        Dim response As String
        Try
            Using sr As New StreamReader(INIFileName)
                Dim line As String
                Do
                    line = sr.ReadLine
                    If line Is Nothing Then
                    Else
                        Dim ReadOption As String = Nothing
                        Dim ReadValue As String = Nothing
                        Dim IsValue As Boolean = False

                        For Each c1 As Char In line
                            If IsValue = False Then
                                If c1 = "=" Then
                                    IsValue = True
                                    c1 = Nothing
                                Else
                                    ReadOption = ReadOption + c1
                                    IsValue = False
                                    If INIOption = INIOption Then
                                        c1 = Nothing
                                        IsValue = True
                                    End If
                                End If
                            Else
                                If c1 = "=" Then
                                    c1 = Nothing
                                Else
                                    ReadValue = ReadValue + c1
                                End If
                            End If
                        Next



                        If INIOption = ReadOption Then
                            'MsgBox(linevalue)
                            response = ReadValue
                        End If
                    End If
                Loop Until line Is Nothing
                sr.Close()

                Return response
            End Using
        Catch ex As Exception
            'MsgBox("The file could not be read:")
            MsgBox(ex.Message)
        End Try
    End Function

End Class
