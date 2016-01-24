Imports System.ComponentModel 'Automatically added by Visual Studio when creating the ContexMenuStrip event
Imports System.Data.SQLite

Public Class Form1

#Region "Enums"
    Enum ChangeIndex
        PrimayKey
        ColName
        ChangeType
        OriginalValue
        NewValue
    End Enum

    Enum ChangeType
        ModfiyVaule
        InsertRow
        DeleteRow
    End Enum

    Enum DGVTag
        PrimaryKeyColIndex
    End Enum
#End Region

#Region "Variables"
    Dim Changes As New ArrayList
    Dim TableName As String
    Dim AppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Sqlite Browser\"
    Dim ChangedValueColor = Color.Yellow, DeletedRowColor = Color.Red
#End Region

#Region "DataGridView Code"
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Dim DataPK = DataGridView1.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyColIndex)), e.RowIndex).Value
        Dim ColName = DataGridView1.Columns.Item(e.ColumnIndex).Name
        If ChangesContainsPKAndColName(DataPK, ColName) = -1 Then
            Dim TempArray(System.Enum.GetValues(GetType(ChangeIndex)).Length - 1)
            TempArray(ChangeIndex.PrimayKey) = DataPK
            TempArray(ChangeIndex.ColName) = ColName
            TempArray(ChangeIndex.ChangeType) = ChangeType.ModfiyVaule
            TempArray(ChangeIndex.OriginalValue) = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            Changes.Add(TempArray)
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim DataPK = DataGridView1.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyColIndex)), e.RowIndex).Value
        Dim ColName = DataGridView1.Columns.Item(e.ColumnIndex).Name
        Dim Index = ChangesContainsPKAndColName(DataPK, ColName)
        Dim TempArray = Changes(Index)
        If TempArray(ChangeIndex.ChangeType) = ChangeType.ModfiyVaule Then
            Dim NewValue = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            If IsNothing(NewValue) Then NewValue = ""
            NewValue = NewValue.ToString.Trim
            If TempArray(ChangeIndex.OriginalValue) <> NewValue Then
                'set/update the new value to be written later
                Changes.Item(Index)(ChangeIndex.NewValue) = NewValue
                DataGridView1.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = ChangedValueColor
                DataGridView1.Item(e.ColumnIndex, e.RowIndex).ToolTipText = "This value was changed. It was originally:" & vbCrLf & TempArray(ChangeIndex.OriginalValue)
            Else
                'if the original value is the same as the changed value, remove it from the changes array
                Changes.RemoveAt(Index)
                DataGridView1.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = DataGridView1.DefaultCellStyle.BackColor
                DataGridView1.Item(e.ColumnIndex, e.RowIndex).ToolTipText = ""
            End If
        End If
    End Sub
#End Region

#Region "Buttons"
    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        SaveChanges(False, True) 'this is the only place where I need to clear out the deleted rows so far
    End Sub

    Private Sub ButtonLoadDB_Click(sender As Object, e As EventArgs) Handles ButtonLoadDB.Click
        If IO.File.Exists(TextBoxDBPath.Text) Then
            SaveCookie()
            If Changes.Count > 0 Then SaveChanges(True)
            ComboBoxTables.Items.Clear()
            For Each i In GetTablesInDataBase(TextBoxDBPath.Text)
                ComboBoxTables.Items.Add(i)
            Next
            If ComboBoxTables.Items.Count > 0 Then
                ComboBoxTables.SelectedIndex = 0
            Else
                MessageBox.Show("No tables found in the database", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("Could not find the file path specificed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
#End Region

#Region "SQL Code"
    Private Function GetColumnsInTable(DataBasePath As String, TableName As String) As String()
        Dim DataTable = New DataTable, ToReturn() As String = Nothing

        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            Using MyCommand As New SQLiteCommand("PRAGMA table_info('" & TableName & "');", MyConnection)
                Using SQLDataAdapter As New SQLite.SQLiteDataAdapter
                    SQLDataAdapter.SelectCommand = MyCommand
                    MyConnection.Open()
                    SQLDataAdapter.Fill(DataTable)
                    MyConnection.Close()
                    ReDim ToReturn(DataTable.Rows.Count - 1)
                    For i = 0 To DataTable.Rows.Count - 1
                        'MessageBox.Show(DataTable.Rows(i)("Name"))
                        'MessageBox.Show(DataTable.Rows(i)("pk"))
                        Dim PrimaryKeyIndex = DataTable.Rows(i)("pk")
                        If IsNumeric(PrimaryKeyIndex) Then
                            If PrimaryKeyIndex > 0 Then
                                If IsNothing(DataGridView1.Tag(DGVTag.PrimaryKeyColIndex)) Then 'need to check for null first before checking the value
                                    DataGridView1.Tag(DGVTag.PrimaryKeyColIndex) = PrimaryKeyIndex - 1
                                ElseIf DataGridView1.Tag(DGVTag.PrimaryKeyColIndex) <> PrimaryKeyIndex - 1 Then
                                    DataGridView1.Tag(DGVTag.PrimaryKeyColIndex) = PrimaryKeyIndex - 1
                                End If
                            End If
                        End If
                        ToReturn(i) = DataTable.Rows(i)("Name").ToString
                    Next
                End Using
            End Using
        End Using

        Return ToReturn
    End Function

    Private Function GetTablesInDataBase(DataBasePath As String)
        Dim ToReturn As Object = Nothing
        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            MyConnection.Open()

            Dim Schema = MyConnection.GetSchema(SQLiteMetaDataCollectionNames.Tables)
            ReDim ToReturn(Schema.Rows.Count - 1)
            For i = 0 To Schema.Rows.Count - 1
                ToReturn(i) = Schema.Rows(i) !TABLE_NAME
                ' Messagebox.Show(Schema.Rows(i) !TABLE_NAME)
                'If Not Schema.Rows(i) !TABLE_NAME = "sqlite_sequence" Then TableName.Items.Add(Schema.Rows(i) !TABLE_NAME)
            Next

            MyConnection.Close()
        End Using
        Return ToReturn
    End Function

    Private Function GetDataBaseRows(DataBasePath As String, TableName As String) As ArrayList
        Dim ToReturn As New ArrayList
        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            Using MyCommand As New SQLiteCommand("SELECT * FROM " & TableName, MyConnection) 'WHERE id = '0'")
                MyConnection.Open()
                Using SQLreader As SQLiteDataReader = MyCommand.ExecuteReader()
                    Dim Cols = GetColumnsInTable(DataBasePath, TableName)
                    While SQLreader.Read()
                        Dim TempArray(Cols.Length - 1) As String
                        For i = 0 To Cols.Length - 1
                            Dim Read = SQLreader(Cols(i))
                            If Read.Equals(DBNull.Value) Then
                                TempArray(i) = ""
                            Else
                                TempArray(i) = Read
                            End If
                        Next
                        ToReturn.Add(TempArray)
                    End While
                End Using
            End Using
            MyConnection.Close()
        End Using
        Return ToReturn
    End Function

    Private Sub UpdateFields(DataBasePath As String, PrimayKeyColName As String, DataArrayList As ArrayList)
        'The Data array needs to be a jagged array (arrays nested in an array)
        'The first index needs to be an array containing the column names
        'The second index needs to be an array containing the values to be updated in the same order as the column names
        'The third index needs to be the ID of the rows to be updated
        'Note, only 1 row (tied to 1 primary key) can be updated at a time.

        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            Dim Commands As New ArrayList
            For Each TheData In DataArrayList
                Dim SetInfo = ""
                Dim MyCommand As New SQLiteCommand(MyConnection)
                If TheData.Length = 3 Then 'Checks to see if the array is three indexes long (see above for structure)
                    If TheData(0).Length = TheData(1).Length Then 'Makes sure the cols array and the values array are the same length (doesn't check to see if they are in the right order though. This must be done before being passed)
                        For i = 0 To TheData(0).Length - 1 'Goes through the length of the two arrays
                            SetInfo += TheData(0)(i) & " = @parama" & i & ", " 'combines the column name to a parameter name
                            MyCommand.Parameters.Add("@parama" & i, ReturnDBType(VarType(TheData(1)(i)))).Value = TheData(1)(i) 'sets the value to the parameter name and gets the data type and set's that as well
                        Next

                        If SetInfo.EndsWith(", ") Then SetInfo = SetInfo.Substring(0, SetInfo.LastIndexOf(", ")) 'removes the extra delimiter
                    Else
                        MessageBox.Show("Something went terribly wrong!", "Error Updating Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                Else
                    MessageBox.Show("Something went terribly wrong!", "Error Updating Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                MyCommand.Parameters.Add("@PrimayKey", DbType.Int32).Value = TheData(2) 'Adds the final parameter
                MyCommand.CommandText = "UPDATE " & TableName & " SET " & SetInfo & " WHERE " & PrimayKeyColName & " = @PrimayKey" 'sets the command text
                Commands.Add(MyCommand) 'adds the command to an array to be batch processed later
            Next

            MyConnection.Open()
            For Each TheCommand As SQLiteCommand In Commands 'batch processes the commands
                TheCommand.ExecuteNonQuery()
                TheCommand.Dispose()
            Next
            MyConnection.Close()
        End Using
    End Sub

    Private Function ReturnDBType(Type As VariantType)
        Select Case Type
            Case VariantType.Array
                MessageBox.Show("The array type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Boolean
                Return DbType.Boolean
            Case VariantType.Byte
                Return DbType.Byte
            Case VariantType.Char
                MessageBox.Show("The char type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Currency
                Return DbType.Currency
            Case VariantType.DataObject
                MessageBox.Show("The data object type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Date
                Return DbType.DateTime
            Case VariantType.Decimal
                Return DbType.Decimal
            Case VariantType.Double
                Return DbType.Double
            Case VariantType.Empty
                MessageBox.Show("The empty type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Error
                MessageBox.Show("The error type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Integer
                Return DbType.Int32
            Case VariantType.Long
                Return DbType.Int64
            Case VariantType.Null
                MessageBox.Show("The null type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Object
                Return DbType.Object
            Case VariantType.Short
                Return DbType.Int16
            Case VariantType.Single
                Return DbType.Single
            Case VariantType.String
                Return DbType.String
            Case VariantType.UserDefinedType
                MessageBox.Show("The user defined type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case VariantType.Variant
                MessageBox.Show("The variant type is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Case Else
                MessageBox.Show("The unknown types is not supported", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
        Return Nothing
    End Function

    Private Sub DeleteRow(DataBasePath As String, ID As Integer)
        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            Using MyCommand As New SQLiteCommand(MyConnection)
                MyCommand.CommandText = "DELETE FROM " & TableName & " WHERE id = @id;"
                MyCommand.Parameters.Add("@id", DbType.Int32).Value = ID
                MyConnection.Open()
                MyCommand.ExecuteNonQuery()
                MyConnection.Close()
            End Using
        End Using
    End Sub
#End Region

#Region "Loading/Closing Stuff"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Sets up the DataGridView
            SetupDGView()

            'Reads the last file that existed when loading the database. If done right it will auto-setup your last database file path
            ReadCookie()

            'Debug mode (True = True equals on)
            If True = False Then
                ListBoxDebugger.Visible = True
                TimerDebugger.Start()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error Loading", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    Private Sub SetupDGView()
        'some properties to setup
        DataGridView1.AllowUserToAddRows = True
        DataGridView1.AllowUserToDeleteRows = True
        DataGridView1.AllowUserToOrderColumns = False
        DataGridView1.AllowUserToResizeColumns = True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Dim Temp(System.Enum.GetValues(GetType(DGVTag)).Length - 1) As Object 'sets up a blank object that will be used an array of any type of data
        DataGridView1.Tag = Temp
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Changes.Count > 0 Then SaveChanges(True)
    End Sub
#End Region

#Region "Cookie Stuff"
    Private Sub ReadCookie()
        If Not IO.Directory.Exists(AppDataDir) Then IO.Directory.CreateDirectory(AppDataDir)
        If IO.File.Exists(AppDataDir & "cookie") Then
            Try
                TextBoxDBPath.Text = IO.File.ReadAllText(AppDataDir & "cookie")
            Catch 'Don't care about errors
            End Try
        End If
    End Sub

    Private Sub SaveCookie()
        If Not IO.Directory.Exists(AppDataDir) Then IO.Directory.CreateDirectory(AppDataDir)
        Try
            IO.File.WriteAllText(AppDataDir & "cookie", TextBoxDBPath.Text)
        Catch 'Don't care about errors
        End Try
    End Sub
#End Region

#Region "Debug Stuff"
    Private Sub TimerDebugger_Tick(sender As Object, e As EventArgs) Handles TimerDebugger.Tick
        Dim DoIt = False
        If ListBoxDebugger.Items.Count <> Changes.Count Then
            DoIt = True
        Else
            For i = 0 To Changes.Count - 1
                If Join(Changes(i), " | ") <> ListBoxDebugger.Items(i) Then
                    DoIt = True
                    Exit For
                End If
            Next
        End If
        If DoIt Then
            ListBoxDebugger.Items.Clear()
            For Each i In Changes
                ListBoxDebugger.Items.Add(Join(i, " | "))
            Next
        End If
    End Sub
#End Region

#Region "Other Program Code"
    'Private Function InsertField(DataBasePath As String, TheDate As String, Amount As Double, Store As String, Category As String, Optional Comment As String = "")
    '    Dim NewID As Integer = -1
    '    Dim MyConnection As New SQLite.SQLiteConnection()
    '    Dim MyCommand As SQLiteCommand
    '    MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
    '    MyConnection.Open()
    '    MyCommand = MyConnection.CreateCommand
    '    ''MyCommand.CommandText = "INSERT INTO " & TableName.SelectedItem & "(col" & DBCol.TheDate & ", col" & DBCol.Amount & ", col" & DBCol.Store & ", col" & DBCol.Category & ", col" & DBCol.Comment & ") VALUES (@thedate, @amount, @store, @category, @comment)"
    '    'Sets the meaning of the parameters.
    '    '(@thedate, @amount, @store, @category, @comment)
    '    MyCommand.Parameters.Add("@thedate", DbType.String).Value = TheDate
    '    MyCommand.Parameters.Add("@amount", DbType.Double).Value = Amount
    '    MyCommand.Parameters.Add("@store", DbType.String).Value = Store
    '    MyCommand.Parameters.Add("@category", DbType.String).Value = Category
    '    MyCommand.Parameters.Add("@comment", DbType.String).Value = Comment
    '    'Runs Query
    '    MyCommand.ExecuteNonQuery()

    '    ''MyCommand.CommandText = "SELECT LAST_INSERT_ROWID() FROM " & TableName.SelectedItem
    '    MyCommand.Parameters.Clear()
    '    Dim LastRowId As Integer = MyCommand.ExecuteScalar() 'only finds 1 cell = Scalar
    '    ''MyCommand.CommandText = "SELECT * FROM " & TableName.SelectedItem & " WHERE id=(@lastrowid)"
    '    MyCommand.Parameters.Add("@lastrowid", DbType.Int32).Value = LastRowId
    '    NewID = MyCommand.ExecuteScalar() 'only finds 1 cell = Scalar

    '    MyCommand.Dispose()
    '    MyConnection.Close()
    '    Return NewID
    'End Function

    ''Private Sub CreateTable(DataBasePath As String, TableName As String)
    ''    Dim MyConnection As New SQLite.SQLiteConnection
    ''    Dim MyCommand As SQLiteCommand
    ''    'This will created the database and setup up the table and columns!
    ''    MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
    ''    MyConnection.Open()
    ''    MyCommand = MyConnection.CreateCommand
    ''    'SQL query to Create Table
    ''    '   MyCommand.CommandText = "CREATE TABLE " & TableName & " (id INTEGER PRIMARY KEY AUTOINCREMENT, col" & DBCol.TheDate & " TEXT, col" & DBCol.Amount & " FLOAT, col" & DBCol.Store & " TEXT, col" & DBCol.Category & " TEXT, col" & DBCol.Comment & " TEXT);"
    ''    MyCommand.ExecuteNonQuery()
    ''    MyCommand.Dispose()
    ''    MyConnection.Close()
    ''End Sub
#End Region

#Region "To Be Organized"
    Private Sub ComboBoxTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxTables.SelectedIndexChanged
        If Changes.Count > 0 Then SaveChanges(True) 'The TableName global variable below hasn't updated to the current selection so I can still save changes before everything is cleared
        TableName = ComboBoxTables.SelectedItem 'I need to set a global variable instead of just using the control because I cannot easily capture what the value is before it is changed. Since most of the events are post-change I would not be able to do a save changes before everything was reset. So this global variable will allow me to save changes with the previous selection before everything is cleared out and deleted
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        Dim Cols = GetColumnsInTable(TextBoxDBPath.Text, TableName)
        If Not IsNothing(Cols) Then
            For Each i In Cols
                DataGridView1.Columns.Add(i, i)
            Next
            Dim Data = GetDataBaseRows(TextBoxDBPath.Text, TableName)
            If Data.Count > 0 Then
                For Each i As String() In Data
                    DataGridView1.Rows.Add(i)
                Next
            End If
            DataGridView1.Columns.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyColIndex))).ReadOnly = True
        End If

    End Sub

    Private Sub SaveChanges(Prompt As Boolean, Optional RemoveDeletedRowsFromDGV As Boolean = False)
        If Prompt Then
            If MessageBox.Show("There are unsaved changes, would you like to save them?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.No Then
                Changes.Clear() 'delete the changes and then exit
                Exit Sub
            End If
        End If
        Changes.Sort(New FirstElementJaggedComparer) 'A custom sorting class that sorts the first elements in the array and tries to convert them to doubles for better sorting
        Dim GroupedChanges, TempCol, TempValue As New ArrayList 'All used to group together the changes with the same ID to send fewer SQL update commands
        Dim LastPK = ""

        For Each i In Changes
            Select Case i(ChangeIndex.ChangeType)
                Case ChangeType.ModfiyVaule
                    If i(ChangeIndex.PrimayKey) = LastPK Then
                        'since the PK is the same, group these together and add them to the temp arrays
                        TempCol.Add(i(ChangeIndex.ColName))
                        TempValue.Add(i(ChangeIndex.NewValue))
                    Else
                        'Adds the grouped together values with the primary key to a jagged array. The first time it will not do this so an if statement prevents errors
                        If TempCol.Count > 0 Then 'to do sort these so they are in the order of the columns in the database (if it is more efficent with sqlite)
                            GroupedChanges.Add({TempCol.ToArray, TempValue.ToArray, LastPK})
                        End If

                        'Resets the temp array lists for the new group
                        TempCol.Clear()
                        TempValue.Clear()
                        LastPK = i(ChangeIndex.PrimayKey) 'updates the last PK to be the current one
                        'Adds the current item to the temp arrays
                        TempCol.Add(i(ChangeIndex.ColName))
                        TempValue.Add(i(ChangeIndex.NewValue))
                    End If
                Case ChangeType.DeleteRow
                    DeleteRow(TextBoxDBPath.Text, i(ChangeIndex.PrimayKey))
                Case ChangeType.InsertRow
                    'To Do
            End Select
        Next

        If TempCol.Count > 0 Then 'This is for when making changes to the fields. This will not run unless there was something changed. This make my changes logic a lot simpler to write above and is why this is out here instead of inside of the loop
            GroupedChanges.Add({TempCol.ToArray, TempValue.ToArray, LastPK})
            TempCol.Clear()
            TempValue.Clear()

            'This will process all of the GroupChanges
            UpdateFields(TextBoxDBPath.Text, DataGridView1.Columns.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyColIndex))).Name, GroupedChanges)
        End If

        Changes.Clear()

        If RemoveDeletedRowsFromDGV Then
            For i = 0 To DataGridView1.Rows.Count - 1
                If i < DataGridView1.Rows.Count Then 'due to MS and their awesome coding, the DataGridView1.Rows.Count doesn't update in the For Loop conditions meaning it will cause an error without an additional check. I could just do an infinite loop and have an exit statement but this will create the integer and automatically add 1 to it so it does save on some coding
                    If DataGridView1.Rows(i).DefaultCellStyle.BackColor = DeletedRowColor Then
                        DataGridView1.Rows.RemoveAt(i)
                        i -= 1
                    End If
                End If
            Next
        End If
    End Sub

    Private Function ChangesContainsPKAndColName(PrimayKey As Int32, ColName As String) As Int32
        For i = 0 To Changes.Count - 1
            If Changes(i)(ChangeIndex.ChangeType) = ChangeType.ModfiyVaule Then
                If Changes(i)(ChangeIndex.PrimayKey) = PrimayKey And Changes(i)(ChangeIndex.ColName).Contains(ColName) Then Return i
            End If
        Next
        Return -1
    End Function

    Private Sub DeleteRowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRowToolStripMenuItem.Click
        For Each Row As DataGridViewRow In DataGridView1.SelectedRows
            'Create a temp array and then store only the necessary information
            Dim TempArray(System.Enum.GetValues(GetType(ChangeIndex)).Length - 1)
            TempArray(ChangeIndex.PrimayKey) = DataGridView1.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyColIndex)), Row.Index).Value
            TempArray(ChangeIndex.ChangeType) = ChangeType.DeleteRow

            'Delete all ModifyValue changes that match the same PK since the row is going to be delete. This should also prevent any bugs in the future (e.g. trying to process a modify value change after the row delete change)
            For i = 0 To Changes.Count - 1
                If i < Changes.Count Then 'due to MS and their awesome coding, the Changes.Count doesn't update in the For Loop conditions meaning it will cause an error without an additional check. I could just do an infinite loop and have an exit statement but this will create the integer and automatically add 1 to it so it does save on some coding
                    If Changes(i)(ChangeIndex.ChangeType) = ChangeType.ModfiyVaule Then
                        If Changes(i)(ChangeIndex.PrimayKey) = TempArray(ChangeIndex.PrimayKey) Then
                            'If the changes arraylist contains the same PK for a ModifyValue change when I am going to delete the row, remove it
                            Changes.RemoveAt(i)
                            i -= 1
                        End If
                    End If
                End If

            Next

            Changes.Add(TempArray)
            ' DataGridView1.Rows.RemoveAt(Row.Index)
            Row.DefaultCellStyle.BackColor = DeletedRowColor
            Row.DefaultCellStyle.SelectionBackColor = Color.Transparent 'disables the selection color
            Row.DefaultCellStyle.SelectionForeColor = Color.Transparent 'disables the selection color
            Row.Selected = False  'since it doesn't update until you unselect it manually I added this
            For Each Cell As DataGridViewCell In Row.Cells
                Cell.ToolTipText = "This row was deleted"
            Next
            Row.ReadOnly = True
        Next
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip1.Opening
        If DataGridView1.SelectedRows.Count = 1 Then
            DeleteRowToolStripMenuItem.Text = "Delete Row"
        ElseIf DataGridView1.SelectedRows.Count > 1 Then
            DeleteRowToolStripMenuItem.Text = "Delete " & DataGridView1.SelectedRows.Count & " Rows"
        Else
            e.Cancel = True
        End If
    End Sub
#End Region
End Class
Class FirstElementJaggedComparer 'A custom sorting class that sorts the first elements in the array and tries to convert them to doubles for better sorting
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim Sort1 = x(Form1.ChangeIndex.PrimayKey), Sort2 = y(Form1.ChangeIndex.PrimayKey)
        'PK can be a string or an int or a double. So I will just convert them to a double since I do not know what type of number they are from the IsNumeric function
        If IsNumeric(Sort1) Then Sort1 = CDbl(Sort1)
        If IsNumeric(Sort2) Then Sort2 = CDbl(Sort2)
        Return Sort1.CompareTo(Sort2)
    End Function
End Class