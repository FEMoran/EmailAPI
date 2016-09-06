Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Data.Common
Imports System.Threading
Imports System.Security.Permissions
Imports Microsoft.Win32
Imports System.Web.Mail
Imports Microsoft.VisualBasic
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Text
Imports System
Imports System.Net
Imports System.Net.Mime
Imports System.ComponentModel

Public Class Form1

    'This is test only.

    Dim QueryStringALARMCOM_ID As Integer
    Dim strdealer As String = ""
    Dim strcsno As String = ""
    Dim strid As Integer = 0

    Dim myReader As SqlDataReader
    Dim mySQLCOnn As SqlConnection
    Dim SQL_Results As DataSet = New DataSet()
    Dim ConnectionStr As String = "Data Source=GC-IIS\SQLEXPRESS;Initial Catalog=NST;user id=website;Password=website@2016;application name= IntranetModule.net;Connect Timeout=240;"
    Dim BolDebug As Boolean
    Dim dtStartDate As DateTime = Now()
    Dim BolSuccess As Boolean = False
    Dim ds As New DataSet
    Dim SQL_Error As String = "0"
    Dim targetDirectory As String = "c:\Daily_Report"

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Vsql As String
        Dim strerror As String = "Sucessfull"
        Dim vcompany As String = ""
        Dim errormessage As String = ""

        GetData()
        'SendData()

       
        Me.Close()

    End Sub


    Sub GetData()


        Dim excelApp As New Excel.Application
        Dim excelBook As Excel.Workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
        Dim partySheet As Excel.Worksheet = Nothing
        Dim newfilename As String
        Dim strdate As String
        strdate = Month(DateAdd(DateInterval.Day, -1, Now())).ToString() + "_" + Day(DateAdd(DateInterval.Day, -1, Now())).ToString() + "_" + Year(DateAdd(DateInterval.Day, -1, Now())).ToString()

        Try
           
            excelBook = excelApp.Workbooks.Open("C:\Daily_Report\Daily_IncomeLog.xls")
            newfilename = "C:\Daily_Report\Archive\Daily_IncomeLog_" & LTrim(RTrim(strdate.ToString())) & ".xls"

            'Delete the file first
            Dim aFile As String
            aFile = newfilename
            If Len(Dir$(aFile)) > 0 Then
                Kill(aFile)
            End If

            excelBook.SaveAs(newfilename)
            partySheet = excelBook.Worksheets("Summary")
            excelBook.Save()

            Dim Vsql As String
            Dim strerror As String = "Sucessfull"
            Dim vcompany As String = ""
            Dim errormessage As String = ""
            Vsql = "Exec dbo.[SP_Get_DailyService_ByCat] "
            ds = GetSQLData(Vsql)

            With partySheet


                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(2)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    k += 1
                Next

                .Range("A" & k - 1 & ":" & "A" & k - 1, "B" & k - 1 & ":" & "B" & k - 1).Merge()
                .Range("A" & k - 1 & ":" & "A" & k - 1, "B" & k - 1 & ":" & "B" & k - 1).HorizontalAlignment = Excel.Constants.xlRight
                .Range("K" & k - 1 & ":" & "K" & k - 1, "K" & k - 1 & ":" & "K" & k - 1).EntireRow.Font.Bold = True
                .Range("I5:I5", "J5:J5").Value = .Range("B9:B9", "B9:B9").Value

                '.Range(k + 1 & ":" & k + 1, "74000:74000").Delete()

            End With

            excelBook.Save()


            partySheet = excelBook.Worksheets("Summary")
            Vsql = "Exec dbo.[SP_Get_DailyService_ByPType] "
            ds = GetSQLData(Vsql)

            With partySheet


                Dim k As Integer = 27
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(2)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)


                    k += 1
                Next

                .Range("A" & k - 1 & ":" & "A" & k - 1, "B" & k - 1 & ":" & "B" & k - 1).Merge()
                .Range("A" & k - 1 & ":" & "A" & k - 1, "B" & k - 1 & ":" & "B" & k - 1).HorizontalAlignment = Excel.Constants.xlRight
                .Range("K" & k - 1 & ":" & "K" & k - 1, "K" & k - 1 & ":" & "K" & k - 1).EntireRow.Font.Bold = True
              
                '.Range(k + 1 & ":" & k + 1, "74000:74000").Delete()

            End With

            excelBook.Save()


            partySheet = excelBook.Worksheets("Gretna")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 1"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()

            partySheet = excelBook.Worksheets("Harahan")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 2"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()

            partySheet = excelBook.Worksheets("Houma")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 3"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()

            partySheet = excelBook.Worksheets("Kenner")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 4"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()

            partySheet = excelBook.Worksheets("Marrero")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 5"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()


            partySheet = excelBook.Worksheets("Metairie")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 6"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()


            partySheet = excelBook.Worksheets("NewOrleans")
            Vsql = "Exec dbo.[SP_Get_DailyService_Details] 7"
            ds = GetSQLData(Vsql)
            With partySheet
                Dim k As Integer = 9
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    .Range("A" & k & ":" & "A" & k, "A" & k & ":" & "A" & k).Value = ds.Tables(0).Rows(row).ItemArray(1)
                    .Range("B" & k & ":" & "B" & k, "B" & k & ":" & "B" & k).Value = ds.Tables(0).Rows(row).ItemArray(3)
                    .Range("C" & k & ":" & "C" & k, "C" & k & ":" & "C" & k).Value = ds.Tables(0).Rows(row).ItemArray(0)
                    .Range("D" & k & ":" & "D" & k, "D" & k & ":" & "D" & k).Value = ds.Tables(0).Rows(row).ItemArray(4)
                    .Range("E" & k & ":" & "E" & k, "E" & k & ":" & "E" & k).Value = ds.Tables(0).Rows(row).ItemArray(5)
                    .Range("F" & k & ":" & "F" & k, "F" & k & ":" & "F" & k).Value = ds.Tables(0).Rows(row).ItemArray(6)
                    .Range("G" & k & ":" & "G" & k, "G" & k & ":" & "G" & k).Value = ds.Tables(0).Rows(row).ItemArray(7)
                    .Range("H" & k & ":" & "H" & k, "H" & k & ":" & "H" & k).Value = ds.Tables(0).Rows(row).ItemArray(8)
                    .Range("I" & k & ":" & "I" & k, "I" & k & ":" & "I" & k).Value = ds.Tables(0).Rows(row).ItemArray(9)
                    .Range("J" & k & ":" & "J" & k, "J" & k & ":" & "J" & k).Value = ds.Tables(0).Rows(row).ItemArray(10)
                    .Range("K" & k & ":" & "K" & k, "K" & k & ":" & "K" & k).Value = ds.Tables(0).Rows(row).ItemArray(11)
                    .Range("L" & k & ":" & "L" & k, "L" & k & ":" & "L" & k).Value = ds.Tables(0).Rows(row).ItemArray(12)
                    .Range("M" & k & ":" & "M" & k, "M" & k & ":" & "M" & k).Value = ds.Tables(0).Rows(row).ItemArray(13)
                    .Range("N" & k & ":" & "N" & k, "N" & k & ":" & "N" & k).Value = ds.Tables(0).Rows(row).ItemArray(14)
                    .Range("O" & k & ":" & "O" & k, "O" & k & ":" & "O" & k).Value = ds.Tables(0).Rows(row).ItemArray(15)
                    .Range("P" & k & ":" & "P" & k, "P" & k & ":" & "P" & k).Value = ds.Tables(0).Rows(row).ItemArray(16)
                    .Range("Q" & k & ":" & "Q" & k, "Q" & k & ":" & "Q" & k).Value = ds.Tables(0).Rows(row).ItemArray(17)
                    .Range("R" & k & ":" & "R" & k, "R" & k & ":" & "R" & k).Value = ds.Tables(0).Rows(row).ItemArray(18)
                    .Range("S" & k & ":" & "S" & k, "S" & k & ":" & "S" & k).Value = ds.Tables(0).Rows(row).ItemArray(19)
                    k += 1
                Next
                .Range("B2:B2", "C2:C2").Value = .Range("C9:C9", "C9:C9").Value
                .Range(k + 1 & ":" & k + 1, "250:250").Delete()

            End With
            excelBook.Save()


        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            'MAKE SURE TO KILL ALL INSTANCES BEFORE QUITING! if you fail to do this
            'The service (excel.exe) will continue to run
            NAR(partySheet)
            excelBook.Close(False)
            NAR(excelBook)
            excelApp.Workbooks.Close()
            NAR(excelApp.Workbooks)
            'quit and dispose app
            excelApp.Quit()
            NAR(excelApp)
            'VERY IMPORTANT
            GC.Collect()
        End Try

    End Sub

    Private Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub

    Function SendData()
        Dim strdate As String
        strdate = Month(DateAdd(DateInterval.Day, -1, Now())).ToString() + "_" + Day(DateAdd(DateInterval.Day, -1, Now())).ToString() + "_" + Year(DateAdd(DateInterval.Day, -1, Now())).ToString()

        Dim filename As String = "C:\Daily_Report\Archive\Daily_IncomeLog_" & LTrim(RTrim(strdate.ToString())) & ".xls"
        Dim Email_Error As String
        Dim Email_ErrorsTo As String = "Reports@notaryshoppe.com"
        Const emailserver As String = "192.168.14.4"
        Const SendEmails As Boolean = True

        Dim StrBody As String = "Please find attached Income Log By Branch For " & strdate.ToString() & " <br/> <br/><br/>"
        StrBody = StrBody & "If you have any questions, Please contact Dhaval Patel @ 847 707-5238. <br/><br/><br/><br/><br/> "
        Dim StrSubject As String = "Daily Income Log For :-  " & strdate.ToString()
        Dim SendMail As New System.Web.Mail.MailMessage()


        Dim attachment As New MailAttachment(filename)
        SendMail.Attachments.Add(attachment)

        SendMail.BodyFormat = System.Web.Mail.MailFormat.Html
        SendMail.From = "Reports@NotaryShoppe.com"
        SendMail.Subject = StrSubject
        SendMail.To = "jlanosga@notaryshoppe.com"
        'SendMail.To = "dpatel2478@hotmail.com"
        SendMail.Bcc = "dpatel2478@hotmail.com"

        SendMail.Body = StrBody
        SendMail.Priority = System.Web.Mail.MailPriority.High
        System.Web.Mail.SmtpMail.SmtpServer = emailserver

        Try

            System.Web.Mail.SmtpMail.Send(SendMail)
            Return Email_Error
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
            Return Email_Error = ex.ToString 'Error Sending Mail
        End Try


    End Function



#Region "Common Functions / SQL / Errors"

    Function GetSQLData(ByVal SQL_Stmt As String)

        Dim SQL_Results As DataSet = New DataSet()
        Dim myConnection As SqlConnection
        SQL_Results.Clear()
        SQL_Results.Reset()
        'HttpContext.Current.Session("SQL_Error") = "0"
        Try
            myConnection = New SqlConnection(ConnectionStr)
            Dim myCommand As New SqlCommand(SQL_Stmt, myConnection)
            myConnection.Open()
            myCommand.CommandTimeout = 30000

            Dim mySqlDataAdapter As New SqlDataAdapter
            mySqlDataAdapter.SelectCommand = myCommand
            mySqlDataAdapter.Fill(SQL_Results)

            myConnection.Close()
            mySqlDataAdapter.Dispose()
            myConnection.Dispose()

        Catch ex As Exception
            SQL_Error = ex.ToString

        End Try

        Return SQL_Results
    End Function

    Function ReadOnlySQLData(ByVal SQL_Stmt As String) As SqlConnection

        Dim myReadOnlyConnection As New SqlConnection(ConnectionStr)

        ' HttpContext.Current.Session("SQL_Error") = 0

        Try
            Dim myCommand As New SqlCommand(SQL_Stmt, myReadOnlyConnection)

            If myReadOnlyConnection.State = ConnectionState.Open Then
                myReadOnlyConnection.Close()
            End If

            myReadOnlyConnection.Open()
            myCommand.CommandTimeout = 300
            myReader = myCommand.ExecuteReader()
            Return myReadOnlyConnection

        Catch e As Exception
            SQL_Error = e.ToString
        End Try

        '    Return myReadOnlyConnection

    End Function

    Function CloseSQL()
        Try
            If myReader Is Nothing Then
                'Do nothing Reader is not fill with any dataset therefore it cannot be closed
            Else
                myReader.Close()
            End If

            mySQLCOnn.Close()
        Catch e As Exception
            Return -1
        End Try

        Return 0
    End Function


#End Region



End Class
