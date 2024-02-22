Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics.Eventing
Imports System.IO
Imports StoredProcedureSSMS.DAL
Imports BObject

Public Class DAL
    Private strConn As String
    Private conn As SqlConnection
    Private cmd As SqlCommand
    Private dr As SqlDataReader

    Public Sub New()
        strConn = "Server=.\SQLEXPRESS02;Database=FP_TEST;Trusted_Connection=True;"
        conn = New SqlConnection(strConn)
    End Sub

    Public Sub CheckCreateAttendanceForAll()
        Try
            Dim strSql = "EXEC CheckCreateAttendanceForAll;"
            conn.Open()
            cmd = New SqlCommand(strSql, conn)
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
            Console.WriteLine("Success")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try

    End Sub

    Public Sub ClockIn(EmployeeID As Integer)
        Try
            Dim strSql = "EXEC ClockInProcedure @EmployeeID;"
            conn.Open()
            cmd = New SqlCommand(strSql, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@EmployeeID", EmployeeID)
            cmd.ExecuteNonQuery()
            Console.WriteLine("Success")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub ClockOut(EmployeeID As Integer)
        Try
            Dim strSql = "EXEC ClockOutProcedure @EmployeeID;"
            conn.Open()
            cmd = New SqlCommand(strSql, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@EmployeeID", EmployeeID)
            cmd.ExecuteNonQuery()
            Console.WriteLine("Success")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub DeleteOldRequest()
        Try
            Dim strSql = "EXEC DeleteOldLeaveRequest;"
            conn.Open()
            cmd = New SqlCommand(strSql, conn)
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
            Console.WriteLine("Success")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub ProcessLeaveRequest(LeaveRequestID As Integer, Status As String)
        Try
            Dim strSql = "EXEC ProcessLeaveRequest @LeaveRequestID, @Status;"
            conn.Open()
            cmd = New SqlCommand(strSql, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@LeaveRequestID", LeaveRequestID)
            cmd.Parameters.AddWithValue("@Status", Status)
            cmd.ExecuteNonQuery()
            Console.WriteLine("Success")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub


    Public Sub RequestLeave(EmployeeID As Integer, LeaveDate As DateTime)

        Try
            Dim strSql = "EXEC RequestLeave @EmployeeID, @LeaveDate;"
            conn.Open()
            cmd = New SqlCommand(strSql, conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@EmployeeID", EmployeeID)
            cmd.Parameters.AddWithValue("@LeaveDate", LeaveDate)
            cmd.ExecuteNonQuery()
            Console.WriteLine("Success")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Function GetAttendance() As List(Of Attendance)
        Dim strSql = "SELECT * FROM Attendance;"
        Dim attendanceList As New List(Of Attendance)()
        conn = New SqlConnection(strConn)
        cmd = New SqlCommand(strSql, conn)
        conn.Open()
        dr = cmd.ExecuteReader()

        While dr.Read()
            Dim attendance As New Attendance()
            attendance.AttendanceID = Convert.ToInt32(dr("AttendanceID"))
            attendance.EmployeeID = If(dr("EmployeeID") Is DBNull.Value, DirectCast(Nothing, Integer?), Convert.ToInt32(dr("EmployeeID")))
            attendance.AttendanceDate = If(dr("AttendanceDate") Is DBNull.Value, DirectCast(Nothing, DateTime?), Convert.ToDateTime(dr("AttendanceDate")))
            attendance.ClockInTime = If(dr("ClockInTime") Is DBNull.Value, DirectCast(Nothing, TimeSpan?), DirectCast(dr("ClockInTime"), TimeSpan))
            attendance.ClockOutTime = If(dr("ClockOutTime") Is DBNull.Value, DirectCast(Nothing, TimeSpan?), DirectCast(dr("ClockOutTime"), TimeSpan))
            attendance.AttendanceStatus = Convert.ToString(dr("AttendanceStatus"))
            attendanceList.Add(attendance)
        End While

        Return attendanceList

    End Function




End Class
