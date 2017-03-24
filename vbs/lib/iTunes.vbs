Option Explicit

Function GetObjFromPerID( sID )
    Set GetObjFromPerID = objItunes.LibraryPlaylist.Tracks.ItemByPersistentID( _
                            Eval( "&H" & Left( sID, 8 ) ), _
                            Eval( "&H" & Right( sID, 8 ) ) _
                        )
End Function

Function GetPerIDFromObj( objTrack )
    GetPerIDFromObj = Right( "0000000" & Hex( objItunes.ITObjectPersistentIDHigh( objTrack ) ), 8 ) & _
                      Right( "0000000" & Hex( objItunes.ITObjectPersistentIDLow( objTrack ) ), 8 )
End Function
