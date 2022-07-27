Attribute VB_Name = "Mng_Excel_Util"
Option Explicit

Public Type T_EXCEL_NEAR_CELL_DATA
    bIsCellDataExist As Boolean
    lRow As Long
    lClm As Long
    sCellValue As String
End Type

Public Function GetNearCellData( _
    ByVal shTrgtSht As Worksheet, _
    ByVal sSrchKey As String, _
    ByVal lRowOffset As Long, _
    ByVal lClmOffset As Long _
) As T_EXCEL_NEAR_CELL_DATA
    Dim rFindResult As Range
    
    Set rFindResult = shTrgtSht.Cells.Find(sSrchKey, LookAt:=xlWhole)
    If rFindResult Is Nothing Then
        GetNearCellData.bIsCellDataExist = False
    Else
        GetNearCellData.bIsCellDataExist = True
        GetNearCellData.lRow = rFindResult.Row + lRowOffset
        GetNearCellData.lClm = rFindResult.Column + lClmOffset
        GetNearCellData.sCellValue = shTrgtSht.Cells( _
                                            GetNearCellData.lRow, _
                                            GetNearCellData.lClm _
                                         ).Value
    End If
End Function
