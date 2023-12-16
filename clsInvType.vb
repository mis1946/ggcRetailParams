﻿'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Inventory Type
'
' Copyright 2012 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  iMac [ 10/19/2016 11:13 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ggcAppDriver

Public Class clsInvType
    Private Const pxeModuleName As String = "clsInvType"
    Private Const pxeTableNamex As String = "Inventory_Type"

    Private p_oApp As GRider
    Private p_oDT As DataTable
    Private p_nEditMode As xeEditMode
    Private p_nRecdStat As xeRecordStat
    Private p_bModified As Boolean
    Private p_bShowMsgx As Boolean

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
                Case 0, 1, 2
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case 0, 1, 2
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select

            RaiseEvent MasterRetreive(Index)
        End Set
    End Property

    Property Master(ByVal Index As String) As Object
        Get
            Select Case Index
                Case "sInvTypCd", "sDescript", "cRecdStat"
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case "sInvTypCd", "sDescript", "cRecdStat"
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
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
        p_oDT.Columns.Add("sInvTypCd", GetType(String)).MaxLength = 4
        p_oDT.Columns.Add("sDescript", GetType(String)).MaxLength = 32
        p_oDT.Columns.Add("cRecdStat", GetType(String)).MaxLength = 1

        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    Function NewRecord() As Boolean
        InitRecord()
        If IsNothing(p_oDT) Then Return False

        p_oDT.Rows.Add()
        p_oDT(0)("sInvTypCd") = ""
        p_oDT(0)("sDescript") = ""
        p_oDT(0)("cRecdStat") = "1"

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

                lsSQL = "UPDATE " & pxeTableNamex & " SET" & _
                                "  sDescript = " & strParm(p_oDT(0)("sDescript")) & _
                                ", cRecdStat = " & strParm(p_oDT(0)("cRecdStat")) & _
                                ", sModified = " & strParm(p_oApp.UserID) & _
                                ", dModified = " & dateParm(p_oApp.SysDate) & _
                        " WHERE sInvTypCd = " & strParm(p_oDT(0)("sInvTypCd"))
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

            lsSQL = "DELETE FROM " & pxeTableNamex & " WHERE sInvTypCd = " & strParm(p_oDT(0)("sInvTypCd"))
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

        loRow = SearchInvType()
        If Not IsNothing(loRow) Then
            InitRecord()
            p_oDT.Rows.Add()
            p_oDT(0)("sInvTypCd") = loRow(0)
            p_oDT(0)("sDescript") = loRow(1)
            p_oDT(0)("cRecdStat") = loRow(2)

            p_nEditMode = xeEditMode.MODE_READY
            Return True
        End If

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        Return False
    End Function

    Function SearchInvType(Optional ByVal lsValue As String = "",
                         Optional ByVal lbCode As Boolean = False) As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        If lbCode Then
            lsSQL = AddCondition(getSQLBrowse, "sInvTypCd = " & strParm(lsValue))
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
                            , "sInvTypCd»sDescript" _
                            , "Code»Descript" _
                            , "" _
                            , "sInvTypCd»sDescript" _
                            , IIf(lbCode, 1, 2))
        End If
    End Function

    Function GetInvType() As DataTable
        Return p_oApp.ExecuteQuery(getSQLBrowse)
    End Function

    Function GetInvType(ByVal lsValue As String,
                        Optional ByVal lbCode As Boolean = False) As DataTable

        If lsValue = "" Then Return Nothing

        If lbCode Then
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sInvTypCd = " & strParm(lsValue)))
        Else
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sDescript = " & strParm(lsValue)))
        End If
    End Function

    Private Function getSQLBrowse() As String
        Return "SELECT" & _
                    "  sInvTypCd" & _
                    ", sDescript" & _
                    ", cRecdStat" & _
                " FROM " & pxeTableNamex & _
                IIf(p_nRecdStat > -1, " WHERE cRecdStat = " & strParm(p_nRecdStat), "")
    End Function

    Private Function isEntryOK() As Boolean
        If p_oDT(0)("sInvTypCd") = "" Or _
            p_oDT(0)("sDescript") = "" Then Return False

        Return True
    End Function
End Class
