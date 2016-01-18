Option Strict Off

Imports System.Data.SQLite
'Make a refrence to C:\Program Files\System.Data.SQLite\2010\bin\System.Data.SQLite.dll
'this is 64bit only due to sqlite being 64bit dll

Public Class Form1

#Region "Enums"
    Enum ChangeIndex
        ID
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
        PrimaryKeyCol
    End Enum
#End Region

#Region "Variables"
    Private Changes As New ArrayList
#End Region

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Sets up the DataGridView
            SetupDGView()

            'Some hard coded stuff for ease of testing
            TextBoxDBPath.Text = "G:\AE\desktop\Database.db3"

            'Debug option
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
                                If IsNothing(DataGridView1.Tag(DGVTag.PrimaryKeyCol)) Then 'need to check for null first before checking the value
                                    DataGridView1.Tag(DGVTag.PrimaryKeyCol) = PrimaryKeyIndex - 1
                                ElseIf DataGridView1.Tag(DGVTag.PrimaryKeyCol) <> PrimaryKeyIndex - 1 Then
                                    DataGridView1.Tag(DGVTag.PrimaryKeyCol) = PrimaryKeyIndex - 1
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
                'If Not Schema.Rows(i) !TABLE_NAME = "sqlite_sequence" Then ComboBoxTable.Items.Add(Schema.Rows(i) !TABLE_NAME)
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

    Private Sub ComboBoxTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxTables.SelectedIndexChanged
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        Dim Cols = GetColumnsInTable(TextBoxDBPath.Text, ComboBoxTables.SelectedItem)
        If Not IsNothing(Cols) Then
            For Each i In Cols
                DataGridView1.Columns.Add(i, i)
            Next
            Dim Data = GetDataBaseRows(TextBoxDBPath.Text, ComboBoxTables.SelectedItem)
            If Data.Count > 0 Then
                For Each i As String() In Data
                    DataGridView1.Rows.Add(i)
                Next
            End If
            DataGridView1.Columns.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyCol))).ReadOnly = True
        End If
    End Sub

    Private Sub ButtonLoadDB_Click(sender As Object, e As EventArgs) Handles ButtonLoadDB.Click
        If IO.File.Exists(TextBoxDBPath.Text) Then
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

    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Dim DGVValue = DataGridView1.Item(CInt(DataGridView1.Tag(DGVTag.PrimaryKeyCol)), e.RowIndex).Value
        If ChangesContainsID(DGVValue) = -1 Then
            Dim TempArray(System.Enum.GetValues(GetType(ChangeIndex)).Length - 1)
            TempArray(ChangeIndex.ID) = DGVValue
            TempArray(ChangeIndex.ChangeType) = ChangeType.ModfiyVaule
            TempArray(ChangeIndex.OriginalValue) = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            Changes.Add(TempArray)
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim DGVValue = DataGridView1.Item(cint(DataGridView1.Tag(DGVTag.PrimaryKeyCol)), e.RowIndex).Value
        Dim Index = ChangesContainsID(DGVValue)
        Dim TempArray = Changes(Index)
        If TempArray(ChangeIndex.ChangeType) = ChangeType.ModfiyVaule Then
            Dim NewValue = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            If IsNothing(NewValue) Then NewValue = ""
            NewValue = NewValue.ToString.Trim
            If TempArray(ChangeIndex.OriginalValue) <> NewValue Then
                'set/update the new value to be written later
                Changes.Item(Index)(ChangeIndex.NewValue) = NewValue
            Else
                'if the original value is the same as the changed value, remove it from the changes array
                Changes.RemoveAt(Index)
            End If
        End If
    End Sub

    Private Function ChangesContainsID(ID As Int32) As Int32
        For i = 0 To Changes.Count - 1
            If Changes(i)(ChangeIndex.ID) = ID Then Return i
        Next
        Return -1
    End Function

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

    Private Sub UpdateFields(DataBasePath As String, IDColName As String, Data As Array)
        'The Data array needs to be a jagged array (arrays nested in an array)
        'The first index needs to be an array containing the column names
        'The second index needs to be an array containing the values to be updated in the same order as the column names
        'The third index needs to be an array containing the IDs of the rows to be updated

        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            Using MyCommand As New SQLiteCommand("", MyConnection)

                Dim ColNames = "", ParamterNames = ""

                If Data.Length = 2 Then
                    If Data(0).Length = Data(1).Length Then
                        For i = 0 To Data(0).Length - 1
                            ColNames += Data(0)(i) & ", "
                            ParamterNames += "@parama" & i & ", "
                            MyCommand.Parameters.Add("@parama" & i, ReturnDBType(VarType(Data(1)(i)))).Value = Data(1)(i)
                        Next

                        If ColNames.EndsWith(", ") Then ColNames = ColNames.Substring(0, ColNames.LastIndexOf(", "))
                        If ParamterNames.EndsWith(", ") Then ParamterNames = ParamterNames.Substring(0, ParamterNames.LastIndexOf(", "))
                    Else
                        MessageBox.Show("Something went terribly wrong!", "Error Updating Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                Else
                    MessageBox.Show("Something went terribly wrong!", "Error Updating Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                ''   MyCommand.Parameters.Add("@id", DbType.Int32).Value = ID

                MyCommand.CommandText = "UPDATE " & ComboBoxTables.SelectedItem & " SET (" & ColNames & ") VALUES (" & ParamterNames & ") WHERE " & IDColName & " = (@id)"
                MyConnection.Open()
                MyCommand.ExecuteNonQuery()
                MyConnection.Close()
            End Using
        End Using
    End Sub

    Private Function ReturnDBType(Type As VariantType)
        Select Case VarType(Type)
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

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        SaveChanges()
    End Sub

    Private Sub SaveChanges()
        For Each i In Changes

        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'This was a great learning experience. However, since I won't be updating all of the same columns for the changes made this will not work
        'I could still try to do this in a single query, but it would require a lot of array maniuplation to get everything aligned properly
        'For example
        '{col1, id2, asdasd}
        '{col2, id432, adsd}
        '{col3, col4, id1,id3,id5, blah}

        'so with that being said, I just need to do a bunch of quereies. I will build the quiereis into an array and execture them while still connected

        Dim IDColName = "id"

        Dim Data As New ArrayList
        Dim ColData = {"col0", "col1", "col2", "col3", "col4"}
        Dim Temp2 = {"Test 1,0 " & Now.ToLongTimeString, "Test 1,1", "Test 1,2", "Test 1,3", "Test 1,4"}
        Data.Add({1, Temp2})
        Temp2 = {"Test 2,0", "Test 2,1 " & Now.ToLongTimeString, "Test 2,2", "Test 2,3", "Test 2,4"}
        Data.Add({2, Temp2})
        Temp2 = {"Test 3,0", "Test 3,1", "Test 3,2 " & Now.ToLongTimeString, "Test 3,3", "Test 3,4"}
        Data.Add({3, Temp2})
        Temp2 = {"Test 4,0", "Test 4,1", "Test 4,2", "Test 4,3 " & Now.ToLongTimeString, "Test 4,4"}
        Data.Add({4, Temp2})
        Temp2 = {"Test 5,0", "Test 5,1", "Test 5,2", "Test 5,3", "Test 5,4 " & Now.ToLongTimeString}
        Data.Add({5, Temp2})

        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & TextBoxDBPath.Text & ";")
            Using MyCommand As New SQLiteCommand("", MyConnection)

                Dim ColNames = "", ParamterNames = "", Command = "", IDs = "", UpdateIDVar = True
                For i = 0 To ColData.Length - 1
                    Command += ColData(i) & " = CASE " & IDColName & " "
                    For Each j In Data
                        If j.Length = 2 Then 'should be an array with 2 items in it (1 int and 1 array)
                            If IsNumeric(j(0)) And j(1).Length = ColData.Length Then 'checks to make sure the ID is a number and the values being updated have the same number of columns
                                Command += "WHEN " & j(0) & " THEN '" & j(1)(i) & "' "
                                If UpdateIDVar Then IDs += j(0) & ", "
                                ''ColNames += i(0)(j) & ", "
                                ''ParamterNames += "@parama" & j & ", "
                                ''MyCommand.Parameters.Add("@parama" & j, ReturnDBType(VarType(i(1)(j)))).Value = i(1)(j)
                                ''If ColNames.EndsWith(", ") Then ColNames = ColNames.Substring(0, ColNames.LastIndexOf(", "))
                                ''If ParamterNames.EndsWith(", ") Then ParamterNames = ParamterNames.Substring(0, ParamterNames.LastIndexOf(", "))
                            Else
                                MessageBox.Show("Something went terribly wrong!", "Error Updating Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Exit Sub
                            End If
                        Else
                            MessageBox.Show("Something went terribly wrong!", "Error Updating Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    Next
                    If UpdateIDVar Then UpdateIDVar = False
                    Command += "END, "
                Next
                If Command.EndsWith(", ") Then Command = Command.Substring(0, Command.LastIndexOf(", "))
                If IDs.EndsWith(", ") Then IDs = IDs.Substring(0, IDs.LastIndexOf(", "))

                Command += " WHERE " & IDColName & " IN (" & IDs & ")"



                ' MyCommand.Parameters.Add("@id", DbType.Int32).Value = ID

                MyCommand.CommandText = "UPDATE " & ComboBoxTables.SelectedItem & " SET " & Command

                MyConnection.Open()
                MyCommand.ExecuteNonQuery()
                MyConnection.Close()
            End Using
        End Using
    End Sub

#Region "Other Program Code"
    'Private Function InsertField(DataBasePath As String, TheDate As String, Amount As Double, Store As String, Category As String, Optional Comment As String = "")
    '    Dim NewID As Integer = -1
    '    Dim MyConnection As New SQLite.SQLiteConnection()
    '    Dim MyCommand As SQLiteCommand
    '    MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
    '    MyConnection.Open()
    '    MyCommand = MyConnection.CreateCommand
    '    ''MyCommand.CommandText = "INSERT INTO " & ComboBoxTable.SelectedItem & "(col" & DBCol.TheDate & ", col" & DBCol.Amount & ", col" & DBCol.Store & ", col" & DBCol.Category & ", col" & DBCol.Comment & ") VALUES (@thedate, @amount, @store, @category, @comment)"
    '    'Sets the meaning of the parameters.
    '    '(@thedate, @amount, @store, @category, @comment)
    '    MyCommand.Parameters.Add("@thedate", DbType.String).Value = TheDate
    '    MyCommand.Parameters.Add("@amount", DbType.Double).Value = Amount
    '    MyCommand.Parameters.Add("@store", DbType.String).Value = Store
    '    MyCommand.Parameters.Add("@category", DbType.String).Value = Category
    '    MyCommand.Parameters.Add("@comment", DbType.String).Value = Comment
    '    'Runs Query
    '    MyCommand.ExecuteNonQuery()

    '    ''MyCommand.CommandText = "SELECT LAST_INSERT_ROWID() FROM " & ComboBoxTable.SelectedItem
    '    MyCommand.Parameters.Clear()
    '    Dim LastRowId As Integer = MyCommand.ExecuteScalar() 'only finds 1 cell = Scalar
    '    ''MyCommand.CommandText = "SELECT * FROM " & ComboBoxTable.SelectedItem & " WHERE id=(@lastrowid)"
    '    MyCommand.Parameters.Add("@lastrowid", DbType.Int32).Value = LastRowId
    '    NewID = MyCommand.ExecuteScalar() 'only finds 1 cell = Scalar

    '    MyCommand.Dispose()
    '    MyConnection.Close()
    '    Return NewID
    'End Function

    ''Private Sub DeleteField(DataBasePath As String, ID As Integer)
    ''    Dim MyConnection As New SQLite.SQLiteConnection()
    ''    Dim MyCommand As SQLiteCommand
    ''    MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
    ''    MyConnection.Open()
    ''    MyCommand = MyConnection.CreateCommand
    ''    'SQL query to Delete Table
    ''    ''MyCommand.CommandText = "DELETE FROM " & ComboBoxTable.SelectedItem & "WHERE id = (@id);"
    ''    MyCommand.Parameters.Add("@id", DbType.Int32).Value = ID
    ''    MyCommand.ExecuteNonQuery()
    ''    MyCommand.Dispose()
    ''    MyConnection.Close()
    ''End Sub

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
End Class
