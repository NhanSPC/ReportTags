Imports pbs.BO
Imports pbs.Helper
Imports System.Globalization

Namespace ReportTags

    Public Class WeekNumberOfMonth
        Inherits FlexCel.Report.TFlexCelUserFunction

        ReadOnly Property Syntax As String
            Get
                Return "<#WeekNumberOfMonth(Date)>"
            End Get
        End Property

        ReadOnly Property Description As String
            Get
                Return "This function return Week number of month. Example: <#WeekNumberOfMonth(07/06/2017)> will return 2 ( 07/06/2017 is the second week of June)"
            End Get
        End Property

        ReadOnly Property Group As String
            Get
                Return ""
            End Get
        End Property

        ReadOnly Property DataStore As String
            Get
                Return ""
            End Get
        End Property


        Public Overrides Function Evaluate(parameters() As Object) As Object
            'get input date
            Dim inputDate As Helper.SmartDate = Helper.SmartDate.Parse(DNz(parameters(0), "T"))

            'caculate
            Dim calendar = New GregorianCalendar

            Dim weekOfYear = calendar.GetWeekOfYear(inputDate.Date, CalendarWeekRule.FirstDay, DayOfWeek.Monday)
            Dim firstWeekOfMonth = calendar.GetWeekOfYear(New DateTime(inputDate.Date.Year, inputDate.Date.Month, 1), CalendarWeekRule.FirstDay, DayOfWeek.Monday)

            Return weekOfYear - firstWeekOfMonth + 1
        End Function
    End Class

End Namespace
