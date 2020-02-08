Attribute VB_Name = "Mng_BasicInfo"
Option Explicit

Private Const SHEET_NAME = "基本情報"

Enum E_ROW
    ROW_KEYWORD_STRT = 21
End Enum

Enum E_CLM
    CLM_REPKEY_SRC = 2
    CLM_REPKEY_DST = 3
End Enum

Type T_REPLACE_INFO
    sRepKeywordSrc As String
    sRepKeywordDst As String
End Type

Type T_BASIC_INFO
    atReplaceInfo() As T_REPLACE_INFO
End Type

Public gtBasicInfo As T_BASIC_INFO

Public Function BasicInfoInit()
    Dim tBasicInfo As T_BASIC_INFO
    gtBasicInfo = tBasicInfo
End Function

Public Function GetBasicInfo()
    Dim lRowIdx As Long
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lRepInfoIdx As Long
    
    With ThisWorkbook.Sheets(SHEET_NAME)
        lStrtRow = ROW_KEYWORD_STRT
        lLastRow = .Cells(.Rows.Count, CLM_REPKEY_SRC).End(xlUp).Row
        
        If lLastRow < lStrtRow Then
            MsgBox "置換元/先の文字列が指定されていません！"
            End
        Else
            'Do Nothing
        End If
        
        lRepInfoIdx = 0
        For lRowIdx = lStrtRow To lLastRow
            If Sgn(gtBasicInfo.atReplaceInfo) = 0 Then
                ReDim Preserve gtBasicInfo.atReplaceInfo(0)
            Else
                ReDim Preserve gtBasicInfo.atReplaceInfo(UBound(gtBasicInfo.atReplaceInfo) + 1)
            End If
            gtBasicInfo.atReplaceInfo(lRepInfoIdx).sRepKeywordSrc = .Cells(lRowIdx, CLM_REPKEY_SRC).Value
            gtBasicInfo.atReplaceInfo(lRepInfoIdx).sRepKeywordDst = .Cells(lRowIdx, CLM_REPKEY_DST).Value
            lRepInfoIdx = lRepInfoIdx + 1
        Next lRowIdx
    End With
    
End Function
