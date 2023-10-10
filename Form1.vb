Imports System.Data.SqlClient
Public Class Form1
    Dim sConn As String = "Data Source=IAU00025L199294\SQLEXPRESS; Initial Catalog=golf; Integrated Security=SSPI;"
    '*** 2023-09-07 you need to instal system.data.sqlClient thorugh NuGet to provide SQL support, though it seems intellisense is oblivious to it
    '*** AND you need to go to project... properties... references... and tick system.data.sqlclient, then intellisense will work
    '*** actually I added Imports system.data.sqlclient and that works

    'https://www.dotnetcurry.com/windows-forms/132/csharp-datagridview-winforms-tutorial
    Private bsource As BindingSource = New BindingSource()
    Private bsCombo As BindingSource = New BindingSource()
    Private ds As DataSet = Nothing
    Private da As SqlDataAdapter

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '*** load the datagridview

        'https://www.codeproject.com/Questions/457084/grid-data-edit-and-update-in-windows-forms
        'https://www.dotnetcurry.com/windows-forms/132/csharp-datagridview-winforms-tutorial

        Dim oConn As New SqlConnection(sConn)
        ds = New DataSet

        '*** this block to populate the in-grid combo
        da = New SqlDataAdapter("SELECT DISTINCT FamilyName AS Fam FROM tblPlayer", oConn)
        da.Fill(ds, "family")
        bsCombo.DataSource = ds.Tables("family")

        Dim cmbCol As DataGridViewComboBoxColumn = DataGridView1.Columns("Column1")
        cmbCol.DataSource = bsCombo
        'these could be set in the designer
        cmbCol.ValueMember = "Fam"
        cmbCol.DisplayMember = "Fam"
        cmbCol.DataPropertyName = "FamilyName"   'this links the selected value to the underlying grid dataset
        '*** the above works, and data is written back to db by the update mechanism
        '*** end in-grid combo.



        da = New SqlDataAdapter("SELECT * FROM tblPlayer", oConn)
        'Dim commandBuilder As SqlCommandBuilder = New SqlCommandBuilder(da)
        'can put the commandBuilder into the update routine
        da.Fill(ds, "players")
        bsource.DataSource = ds.Tables("players")
        DataGridView1.DataSource = bsource






    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'this works, but requires you to buffer the da,dt and call a commandbuilder at the start.
#If False Then
        'my usual approach, does not work
        Dim oConn As New SqlConnection(sConn)
        Dim oDA As New SqlDataAdapter("SELECT * FROM tblPlayer", oConn)
        Dim oDS As New DataSet
        oDA.Fill(oDS, "update")
        Dim builder As SqlCommandBuilder = New SqlCommandBuilder(oDA)

        DataGridView1.BindingContext(oDS.Tables("update")).EndCurrentEdit()
        oDA.Update(oDS.Tables("update"))

#Else
        Dim dt As DataTable = ds.Tables("players")
        Dim commandBuilder As SqlCommandBuilder = New SqlCommandBuilder(da)
        DataGridView1.BindingContext(dt).EndCurrentEdit()
        da.Update(dt)
#End If
    End Sub

    Private Sub DataGridView1_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
        If (Not e.Row.IsNewRow) Then
            Dim res As DialogResult = MessageBox.Show("Are you sure you want to delete this row?", "Delete confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_TabIndexChanged(sender As Object, e As EventArgs) Handles DataGridView1.TabIndexChanged
        'TextBox1.Text = DataGridView1.NewRowIndex
    End Sub

    Private Sub DataGridView1_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.RowEnter
        Try
            TextBox1.Text = e.RowIndex & "   " & DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Catch ex As Exception
            TextBox1.Text = "err"
        End Try


    End Sub
    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        'PART 1: detects changes to combos in the grid
        'but it seems to trigger for every combo even those which are not changed. maybe its triggered when the grid is bound?

        Dim editingComboBox As ComboBox = TryCast(e.Control, ComboBox)
        'gets triggered for each cell you click in, so use trycast instead

        If Not editingComboBox Is Nothing Then
            AddHandler editingComboBox.SelectedIndexChanged, AddressOf editingComboBox_SelectedIndexChanged
        End If
    End Sub

    Private Sub editingComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'PART 2: handler for a change to a grid-combo

        Exit Sub 'too many false triggers

        Dim comboBox1 As ComboBox = CType(sender, ComboBox)
        ' Display index
        MessageBox.Show(comboBox1.SelectedIndex.ToString())
        ' Display value
        MessageBox.Show(comboBox1.Text)
    End Sub


    '*** NOTES ON DataGridView
    '
    'You can define individual columns, their type of control and readonly/visible.  You set the dataPropertyName to the source column
    'there is no validation on these cells and every cell can be edited (unless readonly).  there is a new row at the bottom of
    'the grid that allows a new record to be added
    '
    'does not have DataKeys. Instead you must set up a non-visible column that you can interrogate for the value
    'DataGridView1.Rows(1).Cells(2).Value will bring up the value
    '
    'There is an Autogenerate columns property but not in the designer.  if you define a column it will superceed the auto generated column
    'for a given dataPropertyName.  You could also limit the columns returned in the query or hide columns
    '
    'When displaying data you need to make use of a BindingSource() object and persist this if you later wish to edit/delete/add rows
    'to the undelying dataset
    '
    'SORTING is supported by default on text columns with no extra code required.  just click the column name.
    'note that sorting is not supported on combo columns even though it could sort on dataPropertyName
    '
    'Not clear if column sorting or pagination is supported.  might have to do this by filtering / sorting the dataset
    'YES sorting is supported but I don't know whether this requires a database trip
    '
    'Advanced: populating a combobox in the grid with possible options, selecting current value and triggering an event on new selection
    'Is there an equivalent of rowDatabound?  Nope.  Too hard. so in the meantime, SELECT the row and then have a separate combo and SET
    'button to change the data outside of the grid
    'https://stackoverflow.com/questions/71353293/populate-combobox-in-datagridview-in-vbnet-from-other-table#:~:text=In%20order%20to%20populate%20a,will%20only%20need%20two%20columns.
    '
    '



End Class
