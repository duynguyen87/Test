Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Cryptography


Public Class Form1
    Dim drcom1 As SqlDataReader
    Dim drcom2 As SqlDataReader
    Dim drcom3 As SqlDataReader
    Dim drcom4 As SqlDataReader
    Dim drcom5 As SqlDataReader
    Public dataset As DataTable
    Dim sqlExcute As SQLExcute = New SQLExcute()

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
    End Sub

    'Xem hồ sơ chủ thẻ
    Private Sub btn_ViewPDF_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim test As Date = DateTimePicker2.Text
        'Dim str As String
        'str = Format(test, "ddMMyyyy")
        frm_FileViewer.Show()
        frm_FileViewer.AxAcroPDF1.LoadFile("C:\Users\duynn\Desktop\Gia VLXD Quy II nam 2013 Ha Noi.pdf")
        Dim str As String = Format(DateEdit1.DateTime, "ddMMyyyy")
        MsgBox(str)
    End Sub

    ' Tim kiem thong tin chu the - Chưa validate toDate
    Private Sub btn_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Search.Click
        Dim frDate As String
        Dim toDate As String
        Dim iLoaiChungTu As String
        frDate = Format(DateEdit2.DateTime, "MM/dd/yyyy")
        toDate = Format(DateEdit3.DateTime, "MM/dd/yyyy")
        iLoaiChungTu = ComboBoxEdit2.Text
        Dim SQL As String
        If iLoaiChungTu = String.Empty Then
            SQL = "Select * from DUY_HSCT where [Ngày tạo] >= '" & DateTime.Parse(frDate) & "' and  [Ngày tạo] <=  '" & DateTime.Parse(toDate) & "'"
            GridControl1.DataSource = sqlExcute.GetData(SQL)
        Else
            SQL = "Select * from DUY_HSCT where [Ngày tạo] >= '" & DateTime.Parse(frDate) & "' and  [Ngày tạo] <=   '" & DateTime.Parse(toDate) & "'and [Loại chứng từ] ='" & iLoaiChungTu & "'  "
            GridControl1.DataSource = sqlExcute.GetData(SQL)
        End If

        Dim searchDate As Date
        drcom2 = sqlExcute.Excutereader("Select [Ngày tạo] from DUY_HSCT")
        While (drcom2.Read())
            searchDate = drcom2.GetValue(0)
        End While
    End Sub
    ' Thêm mới hồ sơ chủ thẻ

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        If Trim(TextEdit6.Text) <> String.Empty Then
            Dim Result As MsgBoxResult = MsgBox("Chắn chắn thực hiện ?", MsgBoxStyle.OkCancel)
            If Result = MsgBoxResult.Ok Then

                sqlExcute.CreateCommand(String.Format("Insert into DUY_HSCT Values('{0}',N'{1}',N'{2}',N'{3}',N'{4}','{5}')", TextEdit6.Text, DateEdit1.DateTime))
                GridControl1.DataSource = sqlExcute.GetData("Select * From DUY_HSCT")
                GridControl1.RefreshDataSource()
            End If
        Else : MsgBox("Fill in the blanks")
        End If
    End Sub
    'Tìm đến file hồ sơ mới
    Private Sub btn_BrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_BrowseFile.Click
        Dim openFD As New OpenFileDialog
        Dim strFilename As String = ""
        openFD.InitialDirectory() = "C:\Users\duy nguyen\Desktop"
        openFD.ShowDialog()
        strFilename = openFD.FileName
        openFD.Dispose()
        LinkLabel1.Text = strFilename
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MsgBox(LinkLabel1.Text)
    End Sub
    'Upload file hồ sơ lên file server
    Private Sub btn_UpLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_UpLoad.Click
        Dim filepath As String = "Duy Nguyen"
        Dim filepath_encrypted As String = GenerateHash(filepath)
        MsgBox(filepath_encrypted)
    End Sub


    'Mã hóa Filename 
    Private Function GenerateHash(ByVal SourceText As String) As String
        'Create an encoding object to ensure the encoding standard for the source text
        Dim Ue As New UnicodeEncoding()
        'Retrieve a byte array based on the source text
        Dim ByteSourceText() As Byte = Ue.GetBytes(SourceText)
        'Instantiate an MD5 Provider object
        Dim Md5 As New MD5CryptoServiceProvider()
        'Compute the hash value from the source
        Dim ByteHash() As Byte = Md5.ComputeHash(ByteSourceText)
        'And convert it to String format for return
        Return Convert.ToBase64String(ByteHash)
    End Function


End Class
