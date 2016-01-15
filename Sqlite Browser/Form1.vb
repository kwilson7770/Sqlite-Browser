Imports System.Data.SQLite
'Make a refrence to C:\Program Files\System.Data.SQLite\2010\bin\System.Data.SQLite.dll
'this is 64bit only due to sqlite being 64bit dll

Public Class Form1
    Enum DBCol
        TheDate
        Amount
        Store
        Category
        Comment
    End Enum

    'Some hard coded stuff for testing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Sets up the DataGridView
            SetupDGView()

            'Some hard coded stuff for testing
            TextBoxDBPath.Text = "G:\AE\desktop\Database.db3"
        Catch ex As Exception
            MsgBox(ex.ToString, vbCritical, "Error Loading")
        End Try
    End Sub

    Private Sub SetupDGView()
        'some properties to setup
        DataGridView1.AllowUserToAddRows = True
        DataGridView1.AllowUserToDeleteRows = True
        DataGridView1.AllowUserToOrderColumns = False
        DataGridView1.AllowUserToResizeColumns = True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

    Private Sub CreateTable(DataBasePath As String, TableName As String)
        Dim MyConnection As New SQLite.SQLiteConnection
        Dim MyCommand As SQLiteCommand
        'This will created the database and setup up the table and columns!
        MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
        MyConnection.Open()
        MyCommand = MyConnection.CreateCommand
        'SQL query to Create Table
        MyCommand.CommandText = "CREATE TABLE " & TableName & " (id INTEGER PRIMARY KEY AUTOINCREMENT, col" & DBCol.TheDate & " TEXT, col" & DBCol.Amount & " FLOAT, col" & DBCol.Store & " TEXT, col" & DBCol.Category & " TEXT, col" & DBCol.Comment & " TEXT);"
        MyCommand.ExecuteNonQuery()
        MyCommand.Dispose()
        MyConnection.Close()
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
                        '  MsgBox(DataTable.Rows(i)("Name"))
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
                ' MsgBox(Schema.Rows(i) !TABLE_NAME)
                'If Not Schema.Rows(i) !TABLE_NAME = "sqlite_sequence" Then ComboBoxTable.Items.Add(Schema.Rows(i) !TABLE_NAME)
            Next

            MyConnection.Close()
        End Using
        Return ToReturn
    End Function

    Private Function ReadDataBase(DataBasePath As String, TableName As String) As ArrayList
        Dim ToReturn As New ArrayList
        Using MyConnection As New SQLite.SQLiteConnection("Data Source=" & DataBasePath & ";")
            Using MyCommand As New SQLiteCommand("SELECT * FROM " & TableName, MyConnection) 'WHERE id = '0'")
                MyConnection.Open()
                Using SQLreader As SQLiteDataReader = MyCommand.ExecuteReader()
                    While SQLreader.Read()
                        Dim Cols = GetColumnsInTable(DataBasePath, TableName)
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

    Private Sub UpdateField(DataBasePath As String, ID As Integer, TheDate As String, Amount As Double, Store As String, Category As String, Optional Comment As String = "")
        Dim MyConnection As New SQLite.SQLiteConnection()
        Dim MyCommand As SQLiteCommand
        MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
        MyConnection.Open()
        MyCommand = MyConnection.CreateCommand
        ''MyCommand.CommandText = "UPDATE " & ComboBoxTable.SelectedItem & " SET (col" & DBCol.TheDate & ", col" & DBCol.Amount & ", col" & DBCol.Store & ", col" & DBCol.Category & ", col" & DBCol.Comment & ") VALUES (@thedate, @amount, @store, @category, @comment) WHERE id = (@id)"
        MyCommand.Parameters.Add("@id", DbType.Int32).Value = ID
        MyCommand.Parameters.Add("@thedate", DbType.String).Value = TheDate
        MyCommand.Parameters.Add("@amount", DbType.Double).Value = Amount
        MyCommand.Parameters.Add("@store", DbType.String).Value = Store
        MyCommand.Parameters.Add("@category", DbType.String).Value = Category
        MyCommand.Parameters.Add("@comment", DbType.String).Value = Comment
        MyCommand.ExecuteNonQuery()
        MyCommand.Dispose()
        MyConnection.Close()
    End Sub

    Private Function InsertField(DataBasePath As String, TheDate As String, Amount As Double, Store As String, Category As String, Optional Comment As String = "")
        Dim NewID As Integer = -1
        Dim MyConnection As New SQLite.SQLiteConnection()
        Dim MyCommand As SQLiteCommand
        MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
        MyConnection.Open()
        MyCommand = MyConnection.CreateCommand
        ''MyCommand.CommandText = "INSERT INTO " & ComboBoxTable.SelectedItem & "(col" & DBCol.TheDate & ", col" & DBCol.Amount & ", col" & DBCol.Store & ", col" & DBCol.Category & ", col" & DBCol.Comment & ") VALUES (@thedate, @amount, @store, @category, @comment)"
        'Sets the meaning of the parameters.
        '(@thedate, @amount, @store, @category, @comment)
        MyCommand.Parameters.Add("@thedate", DbType.String).Value = TheDate
        MyCommand.Parameters.Add("@amount", DbType.Double).Value = Amount
        MyCommand.Parameters.Add("@store", DbType.String).Value = Store
        MyCommand.Parameters.Add("@category", DbType.String).Value = Category
        MyCommand.Parameters.Add("@comment", DbType.String).Value = Comment
        'Runs Query
        MyCommand.ExecuteNonQuery()

        ''MyCommand.CommandText = "SELECT LAST_INSERT_ROWID() FROM " & ComboBoxTable.SelectedItem
        MyCommand.Parameters.Clear()
        Dim LastRowId As Integer = MyCommand.ExecuteScalar() 'only finds 1 cell = Scalar
        ''MyCommand.CommandText = "SELECT * FROM " & ComboBoxTable.SelectedItem & " WHERE id=(@lastrowid)"
        MyCommand.Parameters.Add("@lastrowid", DbType.Int32).Value = LastRowId
        NewID = MyCommand.ExecuteScalar() 'only finds 1 cell = Scalar

        MyCommand.Dispose()
        MyConnection.Close()
        Return NewID
    End Function

    Private Sub DeleteField(DataBasePath As String, ID As Integer)
        Dim MyConnection As New SQLite.SQLiteConnection()
        Dim MyCommand As SQLiteCommand
        MyConnection.ConnectionString = "Data Source=" & DataBasePath & ";"
        MyConnection.Open()
        MyCommand = MyConnection.CreateCommand
        'SQL query to Delete Table
        ''MyCommand.CommandText = "DELETE FROM " & ComboBoxTable.SelectedItem & "WHERE id = (@id);"
        MyCommand.Parameters.Add("@id", DbType.Int32).Value = ID
        MyCommand.ExecuteNonQuery()
        MyCommand.Dispose()
        MyConnection.Close()
    End Sub

    Private Sub ComboBoxTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxTables.SelectedIndexChanged
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        Dim Cols = GetColumnsInTable(TextBoxDBPath.Text, ComboBoxTables.SelectedItem)
        If Not IsNothing(Cols) Then
            For Each i In Cols
                DataGridView1.Columns.Add(i, i)
            Next
            Dim Data = ReadDataBase(TextBoxDBPath.Text, ComboBoxTables.SelectedItem)
            If Data.Count > 0 Then
                For Each i As String() In Data
                    DataGridView1.Rows.Add(i)
                Next
            End If
        End If
    End Sub

    Private Sub ButtonLoadDB_Click(sender As Object, e As EventArgs) Handles ButtonLoadDB.Click
        If IO.File.Exists(TextBoxDBPath.Text) Then
            For Each i In GetTablesInDataBase(TextBoxDBPath.Text)
                ComboBoxTables.Items.Add(i)
            Next
            If ComboBoxTables.Items.Count > 0 Then
                ComboBoxTables.SelectedIndex = 0
            Else
                MsgBox("Not tables found in the database", vbInformation, "")
            End If
        Else
            MsgBox("Could not find the file path specificed", vbCritical, "Error")
        End If
    End Sub
End Class
