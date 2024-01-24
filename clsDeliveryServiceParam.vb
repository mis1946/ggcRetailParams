'########################################################################################'
'#        ___          ___          ___           ___       ___                         #'
'#       /\  \        /\  \        /\  \         /\  \     /\  \         ___            #'
'#       \:\  \      /::\  \      /::\  \        \:\  \   /::\  \       /\  \           #'
'#        \:\  \    /:/\:\  \    /:/\:\  \   ___ /::\__\ /:/\:\  \      \:\  \          #'
'#        /::\  \  /::\~\:\  \  /::\~\:\  \ /\  /:/\/__//::\~\:\  \     /::\__\         #'
'#       /:/\:\__\/:/\:\ \:\__\/:/\:\ \:\__\\:\/:/  /  /:/\:\ \:\__\ __/:/\/__/         #'
'#      /:/  \/__/\:\~\:\ \/__/\:\~\:\ \/__/ \::/  /   \:\~\:\ \/__//\/:/  /            #'
'#     /:/  /      \:\ \:\__\   \:\ \:\__\    \/__/     \:\ \:\__\  \::/__/             #'
'#     \/__/        \:\ \/__/    \:\ \/__/               \:\ \/__/   \:\__\             #'
'#                   \:\__\       \:\__\                  \:\__\      \/__/             #'
'#                    \/__/        \/__/                   \/__/                        #'
'#                                                                                      #'
'#                                 DATE CREATED 12-23-2023                              #'
'#                                 DATE LAST MODIFIED 12-26-2023                        #'
'########################################################################################'

Imports MySql.Data.MySqlClient
Imports Microsoft.SqlServer.Server
Imports ggcAppDriver

Public Class clsDeliveryServiceParam
    Private Const pxeModuleName As String = "clsDeliveryServiceParam"
    Private Const pxeTableNamex As String = "Delivery_Service"
    Private Const pxeTableNameHis As String = "Delivery_Service_Charge_History"

    Private p_oApp As GRider
    Private p_oDT As DataTable
    Private p_nEditMode As xeEditMode
    Private p_nRecdStat As xeRecordStat
    Private p_cRiderIDx As String
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
                Case 0, 1, 2, 3, 4, 5, 6
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case 0, 1, 2, 6
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case 3, 5
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
                Case "sRiderIDx", "sBriefDsc", "sDescript", "dPartnerx", "nSrvcChrg", "dSrvcChrg", "cRecdStat"
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case "sRiderIDx", "sBriefDsc", "sDescript", "dPartnerx", "dSrvcChrg", "cRecdStat"
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case "dPartnerx", "dSrvcChrg"
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
                Case 0, 1, 2
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
        p_oDT.Columns.Add("sBriefDsc", GetType(String)).MaxLength = 10
        p_oDT.Columns.Add("sDescript", GetType(String)).MaxLength = 64
        p_oDT.Columns.Add("dPartnerx", GetType(Date))
        p_oDT.Columns.Add("nSrvcChrg", GetType(Decimal))
        p_oDT.Columns.Add("dSrvcChrg", GetType(Date))
        p_oDT.Columns.Add("cRecdStat", GetType(String)).MaxLength = 1


        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    Function NewRecord() As Boolean
        InitRecord()
        If IsNothing(p_oDT) Then Return False

        p_oDT.Rows.Add()
        p_oDT(0)("sRiderIDx") = GetNextCode(pxeTableNamex, "sRiderIDx", False, p_oApp.Connection, False)
        p_oDT(0)("sBriefDsc") = ""
        p_oDT(0)("sDescript") = ""
        p_oDT(0)("dPartnerx") = p_oApp.SysDate
        p_oDT(0)("nSrvcChrg") = "0.0"
        p_oDT(0)("dSrvcChrg") = p_oApp.SysDate
        p_oDT(0)("cRecdStat") = "1"
        
        p_nEditMode = xeEditMode.MODE_ADDNEW
        Return True
    End Function

    Function UpdateRecord() As Boolean
        'p_oTrans = New New_Sales_Order(p_oApp)
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False

        'If Not  Then Return False

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
                'If p_bModified = False Then GoTo endProc
                p_cRiderIDx = p_oDT(0)("sRiderIDx")


                If getPrevData() = False Then
                    GoTo endWithRoll
                End If

                lsSQL = "UPDATE " & pxeTableNamex & " SET" &
                                "  sRiderIDx = " & strParm(p_oDT(0)("sRiderIDx")) &
                                ", sBriefDsc = " & strParm(p_oDT(0)("sBriefDsc")) &
                                ", sDescript = " & strParm(p_oDT(0)("sDescript")) &
                                ", dPartnerx = " & datetimeParm(p_oDT(0)("dPartnerx")) &
                                ", nSrvcChrg = " & strParm(p_oDT(0)("nSrvcChrg")) &
                                ", dSrvcChrg = " & datetimeParm(p_oDT(0)("dSrvcChrg")) &
                                ", cRecdStat = " & strParm(p_oDT(0)("cRecdStat")) &
                                ", sModified = " & strParm(p_oApp.UserID) &
                                ", dModified = " & dateParm(p_oApp.SysDate) &
                                ", dTimeStmp = CURRENT_TIMESTAMP() " &
                        " WHERE sRiderIDx = " & strParm(p_oDT(0)("sRiderIDx"))
            End If



            p_oApp.BeginTransaction()
            If p_oApp.Execute(lsSQL, pxeTableNamex) = 0 Then
                GoTo endWithRoll
            End If
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
        If p_nEditMode <> xeEditMode.MODE_ADDNEW Or p_nEditMode <> xeEditMode.MODE_UPDATE Then
            Return False
        Else

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
        End If
        If p_bShowMsgx Then MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information, "Success")
endProc:
        Return InitRecord()
endWithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Function BrowseRecord() As Boolean
        Dim loRow As DataRow

        loRow = SearchDeliveryService()
        If Not IsNothing(loRow) Then
            InitRecord()
            p_oDT.Rows.Add()
            p_oDT(0)("sRiderIDx") = loRow(0)
            p_oDT(0)("sBriefDsc") = loRow(1)
            p_oDT(0)("sDescript") = loRow(2)
            p_oDT(0)("dPartnerx") = loRow(3)
            p_oDT(0)("nSrvcChrg") = loRow(4)
            p_oDT(0)("dSrvcChrg") = loRow(5)
            p_oDT(0)("cRecdStat") = loRow(6)

            p_nEditMode = xeEditMode.MODE_READY
            Return True
        End If

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        Return False
    End Function

    Function SearchDeliveryService(Optional ByVal lsValue As String = "",
                         Optional ByVal lbCode As Boolean = False) As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        If lbCode Then
            lsSQL = AddCondition(getSQLBrowse, "sRiderIDx = " & strParm(lsValue))
        Else
            lsSQL = AddCondition(getSQLBrowse, "sDescript LIKE " & strParm(lsValue & "%"))
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
                            , "sRiderIDx»sDescript" _
                            , "Code»Description" _
                            , "" _
                            , "sRiderIDx»sDescript" _
                            , IIf(lbCode, 0, 1))
        End If
    End Function

    Function GetDeliveryService() As DataTable
        Return p_oApp.ExecuteQuery(getSQLBrowse)
    End Function

    Function GetDeliveryService(ByVal lsValue As String,
                        Optional ByVal lbCode As Boolean = False) As DataTable

        If lsValue = "" Then Return Nothing

        If lbCode Then
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sRiderIDx = " & strParm(lsValue)))
        Else
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sDescript = " & strParm(lsValue)))
        End If
    End Function

    Private Function getSQLBrowse() As String
        Return "SELECT" &
                    "  sRiderIDx" &
                    ", sBriefDsc" &
                    ", sDescript" &
                    ", dPartnerx" &
                    ", nSrvcChrg" &
                    ", dSrvcChrg" &
                    ", cRecdStat" &
                " FROM " & pxeTableNamex &
                IIf(p_nRecdStat > -1, " WHERE cRecdStat = " & strParm(p_nRecdStat), "")
    End Function

    Private Function getPrevData() As Boolean
        Dim lsSQL As String = ""
        Dim loDT As New DataTable

        Dim id, scharge, sdate As String

        lsSQL = "SELECT" &
                    "  sRiderIDx" &
                    ", nSrvcChrg" &
                    ", dSrvcChrg" &
                " FROM " & pxeTableNamex &
                 " WHERE sRiderIDx = " & strParm(p_cRiderIDx)

        p_oSC.Connection = p_oApp.Connection
        p_oSC.CommandText = lsSQL
        'p_oSC.Parameters.Clear()

        loDT = p_oApp.ExecuteQuery(p_oSC)
        If loDT.Rows.Count > 0 Then
            id = loDT.Rows(0).Item("sRiderIDx")
            scharge = loDT.Rows(0).Item("nSrvcChrg")
            sdate = loDT.Rows(0).Item("dSrvcChrg")

            lsSQL = "SELECT" &
                    "  sRiderIDx" &
                    ", nSrvcChrg" &
                    ", dSrvcChrg" &
                " FROM " & pxeTableNameHis &
                 " WHERE sRiderIDx = " & strParm(p_cRiderIDx) &
                 " AND dSrvcChrg = " & datetimeParm(sdate)

            p_oSC.Connection = p_oApp.Connection
            p_oSC.CommandText = lsSQL
            loDT = p_oApp.ExecuteQuery(p_oSC)

            If loDT.Rows.Count = 0 Then
                lsSQL = "INSERT INTO " & pxeTableNameHis &
                    "  (sRiderIDx" &
                    ", nSrvcChrg" &
                    ", dSrvcChrg" &
                    ", dTimeStmp)" &
                " VALUES (" & strParm(id) &
                    "," & strParm(scharge) &
                    "," & datetimeParm(sdate) &
                    ", CURRENT_TIMESTAMP())"

            Else
                MsgBox("The record has already been modified on the same date.Please contact the MIS Department for further assistance..", MsgBoxStyle.Information, "WARNING!")
                ' lsSQL = "UPDATE " & pxeTableNameHis & " SET" &
                '  "  sRiderIDx = " & strParm(id) &
                '  ", nSrvcChrg = " & strParm(scharge) &
                '  ", dSrvcChrg = " & datetimeParm(sdate) &
                '  ", dTimeStmp = CURRENT_TIMESTAMP() " &
                ' " WHERE sRiderIDx = " & strParm(id)
                Return False
            End If
            p_oSC.Connection = p_oApp.Connection
            p_oSC.CommandText = lsSQL

            loDT = p_oApp.ExecuteQuery(p_oSC)

        End If
        Return True
    End Function

    Private Function isEntryOK() As Boolean
        If p_oDT(0)("sRiderIDx") = "" Or
            p_oDT(0)("sBriefDsc") = "" Or
            p_oDT(0)("sDescript") = "" Then
            MsgBox("Fields seems to be Empty! Please check your entry...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Delivery Service")
            Return False
        End If

        Return True
    End Function

    'Function GetUserApprovals() As Boolean
    '    Dim lofrmUserDisc As New frmUserDisc
    '    Dim loDT As New DataTable

    '    Dim lnCtr As Integer = 0
    '    Dim lbValid As Boolean = False

    '    With lofrmUserDisc
    '        Do
    '            .TopMost = True
    '            .ShowDialog()
    '            If .Cancelled = True Then
    '                Return False
    '            End If

    '            p_oSC.Connection = p_oApp.Connection
    '            p_oSC.CommandText = getSQ_User()
    '            p_oSC.Parameters.Clear()
    '            p_oSC.Parameters.AddWithValue("?sLogNamex", Encrypt(lofrmUserDisc.LogName, xsSignature))
    '            p_oSC.Parameters.AddWithValue("?sPassword", Encrypt(lofrmUserDisc.Password, xsSignature))

    '            loDT = p_oApp.ExecuteQuery(p_oSC)

    '            If loDT.Rows.Count = 0 Then
    '                MsgBox("User Does Not Exist!" & vbCrLf & "Verify log name and/or password.", vbCritical, "Warning")
    '                lnCtr += 1
    '            Else
    '                If Not isUserActive(loDT) Then
    '                    lnCtr = 0
    '                Else
    '                    If loDT.Rows(0).Item("nUserLevl") > xeUserRights.DATAENTRY Then
    '                        lbValid = True
    '                    Else
    '                        MsgBox("User is not allowed to approve this transaction!" & vbCrLf & "Verify user name and/or password.", vbCritical, "Warning")
    '                        lnCtr += 1
    '                    End If
    '                End If
    '            End If
    '        Loop Until lbValid Or lnCtr = 3
    '    End With

    '    If lbValid Then
    '        p_sUserIDxx = loDT.Rows(0).Item("sUserIDxx")
    '        p_sUserName = loDT.Rows(0).Item("sUserName")
    '        p_sLogNamex = loDT.Rows(0).Item("sLogNamex")
    '        p_nUserLevl = loDT.Rows(0).Item("nUserLevl")

    '    End If
    '    Return lbValid
    'End Function
    Private Function isUserActive(ByRef loDT As DataTable) As Boolean
        Dim lnCtr As Integer = 0
        Dim lbMember As Boolean = False

        If loDT.Rows(0).Item("cUserType").Equals(0) Then
            For lnCtr = 0 To loDT.Rows.Count - 1
                If loDT.Rows(0).Item("sProdctID").Equals(p_oApp.ProductID) Then
                    Exit For
                    lbMember = True
                End If
            Next
        Else
            lbMember = True
        End If

        If Not lbMember Then
            MsgBox("User is not a member of this application!!!" & vbCrLf &
               "Application used is not allowed!!!", vbCritical, "Warning")
        End If

        ' check user status
        If loDT.Rows(0).Item("cUserStat").Equals(xeUserStatus.SUSPENDED) Then
            MsgBox("User is currently suspended!!!" & vbCrLf &
                     "Application used is not allowed!!!", vbCritical, "Warning")
            Return False
        End If
        Return True
    End Function

    Private Function getSQ_User() As String
        Return "SELECT sUserIDxx" &
              ", sLogNamex" &
              ", sPassword" &
              ", sUserName" &
              ", nUserLevl" &
              ", cUserType" &
              ", sProdctID" &
              ", cUserStat" &
              ", nSysError" &
              ", cLogStatx" &
              ", cLockStat" &
              ", cAllwLock" &
           " FROM xxxSysUser" &
           " WHERE sLogNamex = ?sLogNamex" &
              " AND sPassword = ?sPassword"
    End Function
End Class
