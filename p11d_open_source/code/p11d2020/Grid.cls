VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private GridArray() As Variant
Private C As TrueDBGrid50.Column
Private WithEvents TrueGrid As TrueDBGrid50.TDBGrid
Attribute TrueGrid.VB_VarHelpID = -1

Public Function InitGrid(grd As TrueDBGrid50.TDBGrid)
  ReDim GridArray(0 To 0)
  
  
  GridArray(0) = ""
  
  Set TrueGrid = grd
  TrueGrid.AllowAddNew = True
  TrueGrid.AllowDelete = True
  
End Function

Private Function MakeBookmark(Index As Long) As String
    ' This support function is used only by the remaining
    ' support functions.  It is not used directly by the
    ' unbound events.  It is a good idea to create a
    ' MakeBookmark function such that all bookmarks can
    ' be created in the same way.  Thus the method by
    ' which bookmarks are created is consistent and easy
    ' to modify.  This function creates a bookmark when
    ' given an array row index.

    ' Since we have data stored in an array, we will just
    ' use the array index as our bookmark.  We will convert
    ' it to a string first, using the Str$ function.  Thus,
    ' if Index = 27, the Bookmark that is created is the
    ' string, " 27".  (Str$ always leaves a leading space
    ' for the sign of the number.)
    
    MakeBookmark = Str$(Index)
End Function

Private Function IndexFromBookmark(Bookmark As Variant, _
        Offset As Long) As Long
    ' This support function is used only by the remaining
    ' support functions.  It is not used directly by the
    ' unbound events.
    
    ' IndexFromBookmark computes the row index that
    ' corresponds to a row that is (Offset) rows from the
    ' row specified by the Bookmark parameter.  For example,
    ' if Bookmark refers to the index 50 of the dataset
    ' array and Offset = -10, then IndexFromBookmark will
    ' return 50 + (-10), or 40.  Thus to get the index of
    ' the row specified by the bookmark itself, call
    ' IndexFromBookmark with an Offset of 0.  If the given
    ' Bookmark is Null, it refers to BOF or EOF.  If
    ' Offset < 0, the grid is requesting rows before the
    ' row specified by Bookmark, and so we must be at EOF
    ' because prior rows do not exist for BOF.  Conversely,
    ' if Offset > 0, we are at BOF.
    
    Dim Index As Long
    
    If IsNull(Bookmark) Then
        If Offset < 0 Then
            ' Bookmark refers to EOF.  Since (MaxRow - 1)
            ' is the index of the last record, we can use
            ' an index of (MaxRow) to represent EOF.
            Index = MaxRow + Offset
        Else
            ' Bookmark refers to BOF.  Since 0 is the index
            ' of the first record, we can use an index of
            ' -1 to represent BOF.
            Index = -1 + Offset
        End If
    Else
        ' Convert string to long integer
        Index = Val(Bookmark) + Offset
    End If
    
    ' Check to see if the row index is valid:
    '   (0 <= Index < MaxRow).
    ' If not, set it to a large negative number to
    ' indicate that it is invalid.
    If Index >= 0 And Index < MaxRow Then
       IndexFromBookmark = Index
    Else
       IndexFromBookmark = -9999
    End If
End Function


Private Function GetNewBookmark() As Variant
    ' GetNewBookmark is called when we need to create a
    ' bookmark for a newly added record (row).

    ' Reserve space for the new row in the array
    ReDim Preserve GridArray(0 To MaxCol - 1, 0 To MaxRow)
    
    ' Since the row index of the last record is
    ' (MaxRow - 1), the next available index for the row
    ' to be added is (MaxRow), so create a bookmark for
    ' the new row using MaxRow as an index.
    GetNewBookmark = MakeBookmark(MaxRow)
    
    ' Increment the number of rows in our dataset
    MaxRow = MaxRow + 1
    
    ' Calibrate the scroll bars based on the new dataset
    ' size.
    TrueGrid.ApproxCount = MaxRow
End Function

' GetRelativeBookmark function
Private Function GetRelativeBookmark(Bookmark As Variant, _
        Offset As Long) As Variant
    ' GetRelativeBookmark is used to get a bookmark for a
    ' row that is a specified number of rows away from the
    ' given row.  Offset specifies the number of rows to
    ' move.  A positive Offset indicates that the desired
    ' row is after the one referred to by Bookmark, and a
    ' negative Offset means it is before the one referred
    ' to by Bookmark.
    
    Dim Index As Long
    
    ' Compute the row index for the desired row
    Index = IndexFromBookmark(Bookmark, Offset)
    If Index < 0 Or Index >= MaxRow Then
        ' Index refers to a row before the first or after
        ' the last, so just return Null.
        GetRelativeBookmark = Null
    Else
        GetRelativeBookmark = MakeBookmark(Index)
    End If
End Function

' GetUserData function
Public Function GetUserData(Bookmark As Variant, _
        Col As Integer) As Variant
    ' In this example, GetUserData is called by
    ' UnboundReadData to ask the user what data should be
    ' displayed in a specific cell in the grid.  The grid
    ' row the cell is in is the one referred to by the
    ' Bookmark parameter, and the column it is in it given
    ' by the Col parameter.  GetUserData is called on a
    ' cell-by-cell basis.
    
    Dim Index As Long
        
    ' Figure out which row the bookmark refers to
    Index = IndexFromBookmark(Bookmark, 0)
    
    If Index < 0 Or Index >= MaxRow Or _
       Col < 0 Or Col >= MaxCol Then
        ' Cell position is invalid, so just return null to
        ' indicate failure
        GetUserData = Null
    Else
        GetUserData = GridArray(Col, Index)
    End If
End Function

' StoreUserData function
Public Function StoreUserData(Bookmark As Variant, _
        Col As Integer, Userval As Variant) As Boolean
    ' StoreUserData is called from UnboundWriteData to
    ' write the newly changed information in the grid to
    ' the array.  This function is called once for each
    ' field that needs to be stored.  The grid cell that
    ' this function is called to store is referenced in
    ' the same way as in GetUserData.

    Dim Index As Long
    
    ' Figure out which row the bookmark refers to
    Index = IndexFromBookmark(Bookmark, 0)
    
    If Index < 0 Or Index >= MaxRow Or Col < 0 Or _
       Col >= MaxCol Then
        ' Cell position is invalid, so just return null
        ' to indicate failure
        StoreUserData = False
    Else
        StoreUserData = True
        GridArray(Col, Index) = Userval
    End If
End Function

' DeleteRow function
Private Function DeleteRow(Bookmark As Variant) As Boolean
    ' DeleteRow is called to logically delete a record
    ' (row) from the array.  The row to be deleted is
    ' given by the Bookmark parameter.

    Dim I As Long, Index As Long
    Dim J As Integer

    ' Figure out which row index the bookmark refers to
    Index = IndexFromBookmark(Bookmark, 0)
    
    If Index < 0 Or Index >= MaxRow Then
        ' Row index is invalid, so exit the delete,
        ' indicating failure.
        DeleteRow = False
        Exit Function
    End If
    
    ' Decrement the number of rows in the dataset
    MaxRow = MaxRow - 1
    
    ' Shift the data in the array to fill the empty space
    ' vacated by the deleted row
    For I = Index To MaxRow - 1
        For J = 0 To MaxCol - 1
            GridArray(J, I) = GridArray(J, I + 1)
        Next J
    Next I
    
    ' Resize array to free storage space used by deleted
    ' row
    If MaxRow > 0 Then
        ' If MaxRow = 0, then (MaxRow - 1) equals -1.
        ' This causes an error in the statement below, so
        ' we handle this special case in the Else clause.
        ReDim Preserve GridArray(0 To MaxCol - 1, 0 _
              To MaxRow - 1)
    Else
        ReDim Preserve GridArray(0 To MaxCol - 1, 0)
    End If
    DeleteRow = True
    
    ' Calibrate scroll bars based on new dataset size
    TrueGrid.ApproxCount = MaxRow
End Function



Private Sub TrueGrid_UnboundAddData(ByVal RowBuf As TrueDBGrid50.RowBuffer, NewRowBookmark As Variant)
  ' UnboundAddData is fired when the user adds a new
    ' row of data to an unbound grid to alert your
    ' application that it must add a new row of data to
    ' its unbound dataset.  This event will NOT be fired
    ' if the grid's AllowAddNew property is set to False.
    
    ' Assume that a function StoreUserData(bookmark, col,
    ' value) takes a row bookmark, a column index, and a
    ' variant with the appropriate data to be stored in an
    ' array or database.  The function returns True if the
    ' data is acceptable and can be stored, False otherwise.
    
    ' First, get a bookmark for the new row.  Do this with
    ' a function, GetNewBookmark.  This function allocates
    ' a new row of data in the storage medium (array or
    ' database), and returns a variant containing a
    ' bookmark for the added row.
        
    Dim NewVal As Variant
    Dim I As Integer
    
    ' Create a new row in the array and get a bookmark for
    ' the new row
    NewRowBookmark = GetNewBookmark()

    ' Loop over all the columns of the row, storing
    ' non-Null values
    For I = 0 To RowBuf.ColumnCount - 1
        ' Note that since only one row of data may be
        ' added at a time, the first index of the row
        ' buffer Value array is always 0.
        NewVal = RowBuf.value(0, I)
        If IsNull(NewVal) Then
            ' The row buffer does not contain a value for
            ' this column.  A default value should be set.
            ' A convenient value is the default value for
            ' the column.
            NewVal = TrueGrid.Columns(I).DefaultValue
        End If
    
        ' Now store the new values.
        If Not StoreUserData(NewRowBookmark, I, NewVal) Then
            ' Storage of the data has failed.  Delete the
            ' added row using a function, DeleteRow, which
            ' takes a bookmark as an argument.  Also,
            ' indicate that the update has failed by setting
            ' RowCount to 0.
            DeleteRow NewRowBookmark
            RowBuf.RowCount = 0
            Exit Function     ' It failed, so exit the event.
        End If
    Next I

End Sub

Private Sub TrueGrid_UnboundDeleteRow(Bookmark As Variant)
  If Not DeleteRow(Bookmark) Then Bookmark = Null
End Sub

Private Sub TrueGrid_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal Offset As Long, ApproximatePosition As Long)
   
        
    Dim ColIndex As Integer, J As Integer
    Dim RowsFetched As Integer, I As Long
    Dim NewPosition As Long, Bookmark As Variant

    RowsFetched = 0
    
    For I = 0 To RowBuf.RowCount - 1
        ' Get the bookmark of the next available row
        Bookmark = GetRelativeBookmark(StartLocation, _
                   Offset + I)
    
        ' If the next row is BOF or EOF, then stop fetching
        ' and return any rows fetched up to this point.
        If IsNull(Bookmark) Then Exit For
    
        ' Place the record data into the row buffer
        For J = 0 To RowBuf.ColumnCount - 1
            ColIndex = RowBuf.ColumnIndex(I, J)
            RowBuf.value(I, J) = GetUserData(Bookmark, _
                                 ColIndex)
        Next J
    
        ' Set the bookmark for the row
        RowBuf.Bookmark(I) = Bookmark
    
        ' Increment the count of fetched rows
        RowsFetched = RowsFetched + 1
    Next I
    
    ' Tell the grid how many rows were fetched
    RowBuf.RowCount = RowsFetched
    
    ' Set the approximate scroll bar position.  Only
    ' nonnegative values of IndexFromBookmark() are valid.
    NewPosition = IndexFromBookmark(StartLocation, Offset)
    If NewPosition >= 0 Then _
       ApproximatePosition = NewPosition

End Sub

Private Sub TrueGrid_UnboundWriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, WriteLocation As Variant)
    ' UnboundWriteData event is fired when the user
    ' modifies an existing row within an unbound grid and
    ' attempts to commit the changes by moving to another
    ' row or calling the Update method of the grid.  This
    ' notifies the program that data has changed in the
    ' grid and it needs to be stored in the data source.
    
    ' RowBuf is the row buffer from which you retrieve the
    ' data to be stored in the data source.  Since only
    ' one row can be updated at a time, RowBuf.RowCount
    ' always equals 1.  If a particular column in the row
    ' has not been changed, its element in the RowBuf.Value
    ' array will be null.
    
    ' WriteLocation is a bookmark that identifies the row
    ' to be updated.
    
    ' Assume that a function StoreUserData(bookmark, col,
    ' value) takes a row bookmark, a column index, and a
    ' variant with the appropriate data to be stored in
    ' an array or database.  The function returns True if
    ' the data is acceptable and can be stored, False
    ' otherwise.
    
    Dim I As Integer
    
    ' Loop over all the columns of the row, storing
    ' non-Null values
    For I = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.value(0, I)) Then
            If Not StoreUserData(WriteLocation, I, _
               RowBuf.value(0, I)) Then
                RowBuf.RowCount = 0 ' Tell the grid the
                                    ' update failed, so
            End If                  ' exit the event.
        End If
    Next I

End Sub
