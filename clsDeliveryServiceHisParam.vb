Imports MySql.Data.MySqlClient
Imports ggcAppDriver
#Disable Warning BC40056 ' Namespace or type specified in Imports statement doesn't contain any public member or cannot be found
Imports ggcRetailSales
#Enable Warning BC40056 ' Namespace or type specified in Imports statement doesn't contain any public member or cannot be found
Imports Microsoft.SqlServer.Server
Public Class clsDeliveryServiceHisParam
    Private Const pxeModuleName As String = "clsDeliveryServiceHisParam"
    Private Const pxeTableNamex As String = "Delivery_Service_Charge_History"

    Private p_oApp As GRider
    Private p_oDT As DataTable
    Private p_nEditMode As xeEditMode
    Private p_nRecdStat As xeRecordStat
    Private p_bModified As Boolean
    Private p_bShowMsgx As Boolean
    Private p_oSC As New MySqlCommand

    Private p_sUserIDxx As String
    Private p_sUserName As String
    Private p_sLogNamex As String
    Private p_nUserLevl As Integer

    Public Event MasterRetreive(ByVal lnIndex As Integer)

    Public Sub New(ByVal foRider As GRider, Optional ByVal fbShowMsg As Boolean = False)
        p_oApp = foRider
        p_nRecdStat = -1
        p_bShowMsgx = fbShowMsg

        InitRecord()
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnRecdStat As xeRecordStat, Optional ByVal fbShowMsg As Boolean = False)
        p_oApp = foRider
        p_nRecdStat = fnRecdStat
        p_bShowMsgx = fbShowMsg

        InitRecord()
    End Sub

    Property Master(ByVal Index As Integer) As Object
        Get
            Select Case Index
                Case 0, 5, 4
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case 0
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case 5
                    If p_oDT(0)(Index) <> CDate(value) Then
                        p_oDT(0)(Index) = CDate(value)
                        p_bModified = True
                    End If
                Case 4
                    p_oDT(0).Item(Index) = value
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select

            RaiseEvent MasterRetreive(Index)
        End Set
    End Property

    Property Master(ByVal Index As String) As Object
        Get
            Select Case Index
                Case "sRiderIDx", "nSrvcChrg", "dSrvcChrg"
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case "sRiderIDx"
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case "dSrvcChrg"
                    If p_oDT(0)(Index) <> CDate(value) Then
                        p_oDT(0)(Index) = CDate(value)
                        p_bModified = True
                    End If

                Case "nSrvcChrg"
                    'p_oDT(0)(Index) = Double.Parse(value)

                    p_oDT(0).Item(Index) = value

                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
            RaiseEvent MasterRetreive(Index)


        End Set
    End Property

    ReadOnly Property MasFldSze(ByVal Index As Integer)
        Get
            Select Case Index
                Case 0, 5, 4
                    MasFldSze = p_oDT.Columns(Index).MaxLength
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
                    Return 0
            End Select
        End Get
    End Property

    ReadOnly Property EditMode()
        Get
            EditMode = p_nEditMode
        End Get
    End Property

    Function InitRecord() As Boolean
        p_oDT = New DataTable
        p_oDT.Columns.Add("sRiderIDx", GetType(String)).MaxLength = 3
        p_oDT.Columns.Add("dSrvcChrg", GetType(Date))
        p_oDT.Columns.Add("nSrvcChrg", GetType(Decimal))
        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    Function NewRecord() As Boolean
        InitRecord()
        If IsNothing(p_oDT) Then Return False

        p_oDT.Rows.Add()
        p_oDT(0)("sRiderIDx") = GetNextCode(pxeTableNamex, "sRiderIDx", False, p_oApp.Connection, False)
        p_oDT(0)("dSrvcChrg") = ""
        p_oDT(0)("nSrvcChrg") = "0.0"

        p_nEditMode = xeEditMode.MODE_ADDNEW
        Return True
    End Function

    Function UpdateRecord() As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False


        If IsNothing(p_oDT) Then Return False
        If p_oDT.Rows.Count = 0 Then Return False

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
    End Function

    Function CancelUpdate() As Boolean
        p_nEditMode = xeEditMode.MODE_READY

        Return True
    End Function

    Function SaveRecord() As Boolean
        Dim lsSQL As String

        Dim lsProcName As String = pxeModuleName & "." & "SaveRecord"

        If p_nEditMode <> xeEditMode.MODE_ADDNEW And p_nEditMode <> xeEditMode.MODE_UPDATE Then Return False

        Try
            If Not isEntryOK() Then Return False

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then

                lsSQL = ADO2SQL(p_oDT, pxeTableNamex, , p_oApp.UserID, p_oApp.SysDate)
            Else
                If p_bModified = False Then GoTo endProc

                lsSQL = "UPDATE " & pxeTableNamex & " SET" &
                                "  sMeasurNm = " & strParm(p_oDT(0)("sMeasurNm")) &
                                ", cRecdStat = " & strParm(p_oDT(0)("cRecdStat")) &
                                ", sModified = " & strParm(p_oApp.UserID) &
                                ", dModified = " & dateParm(p_oApp.SysDate) &
                        " WHERE sMeasurID = " & strParm(p_oDT(0)("sMeasurID"))
            End If

            p_oApp.BeginTransaction()
            If p_oApp.Execute(lsSQL, pxeTableNamex) = 0 Then GoTo endWithRoll
        Catch ex As Exception
            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")
            GoTo endWithRoll
        End Try

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        p_oApp.CommitTransaction()
endProc:
        If p_bShowMsgx Then MsgBox("Record Saved Successfully.", MsgBoxStyle.Information, "Success")
        Return True
endWithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Function DeleteRecord() As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False

        Dim lsSQL As String

        Dim lsProcName As String = pxeModuleName & ".DeleteRecord"

        If IsNothing(p_oDT) Then Return False
        If p_oDT.Rows.Count = 0 Then Return False

        If p_bShowMsgx Then
            If MsgBox("Are you sure to delete this record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                Return False
            End If
        End If

        Try
            p_oApp.BeginTransaction()

            lsSQL = "DELETE FROM " & pxeTableNamex & " WHERE sRiderIDx = " & strParm(p_oDT(0)("sRiderIDx"))
            p_oApp.Execute(lsSQL, pxeTableNamex)
        Catch ex As Exception
            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")
            GoTo endWithRoll
        End Try

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        p_oApp.CommitTransaction()

        If p_bShowMsgx Then MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information, "Success")
endProc:
        Return InitRecord()
endWithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Function BrowseRecord() As Boolean
        Dim loRow As DataRow

        loRow = SearchMeasure()
        If Not IsNothing(loRow) Then
            InitRecord()
            p_oDT.Rows.Add()
            p_oDT(0)("sRiderIDx") = loRow(0)
            p_oDT(0)("dSrvcChrg") = loRow(1)
            p_oDT(0)("nSrvcChrg") = loRow(2)

            p_nEditMode = xeEditMode.MODE_READY
            Return True
        End If

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        Return False
    End Function

    Function SearchMeasure(Optional ByVal lsValue As String = "",
                         Optional ByVal lbCode As Boolean = False) As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        If lbCode Then
            lsSQL = AddCondition(getSQLBrowse, "sRiderIDx = " & strParm(lsValue))
        Else
            lsSQL = AddCondition(getSQLBrowse, "dSrvcChrg LIKE " & strParm(lsValue & "%"))
        End If

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            Return Nothing
        ElseIf loDT.Rows.Count = 1 Then
            Return loDT.Rows(0)
        Else
            Return KwikSearch(p_oApp _
                            , getSQLBrowse _
                            , True _
                            , lsValue _
                            , "sRiderIDx»dSrvcChrg" _
                            , "Code»Description" _
                            , "" _
                            , "sRiderIDx»dSrvcChrg" _
                            , IIf(lbCode, 0, 1))
        End If
    End Function

    Function GetMeasure() As DataTable
        Return p_oApp.ExecuteQuery(getSQLBrowse)
    End Function

    Function GetMeasure(ByVal lsValue As String,
                        Optional ByVal lbCode As Boolean = False) As DataTable

        If lsValue = "" Then Return Nothing

        If lbCode Then
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sRiderIDx = " & strParm(lsValue)))
        Else
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "dSrvcChrg = " & strParm(lsValue)))
        End If
    End Function

    Private Function getSQLBrowse() As String
        Return "SELECT" &
                    "  sRiderIDx" &
                    ", dSrvcChrg" &
                    ", nSrvcChrg" &
                " FROM " & pxeTableNamex &
                IIf(p_nRecdStat > -1, " WHERE cRecdStat = " & strParm(p_nRecdStat), "")
    End Function

    Private Function isEntryOK() As Boolean
        If p_oDT(0)("sRiderIDx") = "" Or
            p_oDT(0)("dSrvcChrg") = "" Then Return False

        Return True
    End Function
End Class
