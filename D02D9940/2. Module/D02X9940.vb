#Region "Khai báo Structure"

''' <summary>
''' Khai báo Structure cho phần Tùy chọn của Module
''' </summary>
Public Structure StructureOption
    ''' <summary>
    ''' Hỏi trước khi lưu
    ''' </summary>
    Public MessageAskBeforeSave As Boolean
    ''' <summary>
    ''' Thông báo khi lưu thành công
    ''' </summary>
    Public MessageWhenSaveOK As Boolean
    ''' <summary>
    ''' Hiển thị form chọn kỳ kế toán khi chạy chương trình
    ''' </summary>
    Public ViewFormPeriodWhenAppRun As Boolean
    ''' <summary>
    ''' Lưu giá trị gần nhất
    ''' </summary>
    Public SaveLastRecent As Boolean
    ''' <summary>
    ''' Lưu đơn vị mặc định
    ''' </summary>
    Public DefaultDivisionID As String
    ''' <summary>
    ''' Khóa thành tiền quy đổi
    ''' </summary>
    Public LockConvertedAmount As Boolean
    ''' <summary>
    ''' Làm tròn thành tiền quy đổi
    ''' </summary>
    Public RoundConvertedAmount As Boolean
    ''' <summary>
    ''' Hiển thị quy trình sơ đồ nghiệp vụ
    ''' </summary>
    Public ViewWorkflow As Boolean
    ''' <summary>
    ''' Ngôn ngữ báo cáo
    ''' </summary>
    Public ReportLanguage As Byte
    ''' <summary>
    ''' Hiển thị đường dẫn báo cáo cho lần sau
    ''' </summary>
    Public ShowReportPath As Boolean
    '------------------------------------------------------------------------
    '  D02 Options here
    '------------------------------------------------------------------------
    Public DepAndEmpID As Boolean 'Bộ phận tiếp nhận và người tiếp nhận
    Public CipName As Boolean 'Tên XDCB
    Public CipDescription As Boolean 'Diễn giải
End Structure


''' <summary>
''' Khai báo structure cho phần Thiết lập hệ thống
''' </summary>
Public Structure StructureSystem
    ''' <summary>
    ''' Đơn vị mặc định
    ''' </summary>
    Public DefaultDivisionID As String
    ''' <summary>
    ''' TK tài sản
    ''' </summary>
    Public DefAssetAccountID As String
    ''' <summary>
    ''' TK khấu hao
    ''' </summary>
    Public DefDepreciationAccountID As String
    ''' <summary>
    ''' Nguồn vốn
    ''' </summary>
    Public DefSourceID As String
    ''' <summary>
    ''' Phân bổ KH
    ''' </summary>
    Public DefAssignmentID As String
    ''' <summary>
    ''' Phương pháp KH
    ''' </summary>
    Public MethodID As String
    ''' <summary>
    ''' Xử lý KH kỳ cuối
    ''' </summary>
    Public MethodEndID As String
    ''' <summary>
    ''' Các bút toán giảm TS
    ''' </summary>
    Public DecreaseAsset As Boolean
    ''' <summary>
    ''' Tạo mã tự động cho tài sản cố định
    ''' </summary>
    Public AssetAuto As Integer

    ''' <summary>
    ''' Phân loại 1
    ''' </summary>
    Public AssetS1Enabled As Boolean
    ''' <summary>
    ''' Phân loại 2
    ''' </summary>
    Public AssetS2Enabled As Boolean
    ''' <summary>
    ''' Phân loại 3
    ''' </summary>
    Public AssetS3Enabled As Boolean
    ''' <summary>
    ''' AssetS1Default
    ''' </summary>
    Public AssetS1Default As String
    ''' <summary>
    ''' AssetS2Default
    ''' </summary>
    Public AssetS2Default As String
    ''' <summary>
    ''' AssetS3Default
    ''' </summary>
    Public AssetS3Default As String
    ''' <summary>
    ''' S1Length
    ''' </summary>
    Public S1Length As Integer
    ''' <summary>
    ''' S2Length
    ''' </summary>
    Public S2Length As Integer
    ''' <summary>
    ''' S3Length
    ''' </summary>
    Public S3Length As Integer
    ''' <summary>
    ''' Dấu phân cách
    ''' </summary>
    Public AssetSeperated As Boolean
    ''' <summary>
    ''' Dấu phân cách
    ''' </summary>
    Public AssetSeperator As String
    ''' <summary>
    ''' Dạng hiển thị
    ''' </summary>
    Public AssetOutputOrder As String
    ''' <summary>
    ''' Độ dài số
    ''' </summary>
    Public AutoNumberLength As Integer
    ''' <summary>
    ''' Độ dài mã
    ''' </summary>
    Public AssetOutputLength As Integer
    ''' <summary>
    ''' Mã TSCĐ và mã CCDC tăng liên tục
    ''' </summary>
    Public IsAssetIDForD02D43 As Boolean
    ''' <summary>
    ''' BĐS đầu tư
    ''' </summary>
    Public UseProperty As Boolean

    'Phan danh cho XDCB
    ''' <summary>
    ''' Độ dài mã Incident 69247
    ''' </summary>
    Public CIPAuto As Integer

    ''' <summary>
    ''' Phân loại 1
    ''' </summary>
    Public CIPS1Enabled As Boolean
    ''' <summary>
    ''' Phân loại 2
    ''' </summary>
    Public CIPS2Enabled As Boolean
    ''' <summary>
    ''' Phân loại 3
    ''' </summary>
    Public CIPS3Enabled As Boolean
    ''' <summary>
    ''' AssetS1Default
    ''' </summary>
    Public CIPS1Default As String
    ''' <summary>
    ''' AssetS2Default
    ''' </summary>
    Public CIPS2Default As String
    ''' <summary>
    ''' AssetS3Default
    ''' </summary>
    Public CIPS3Default As String
    ''' <summary>
    ''' S1Length
    ''' </summary>
    ''' <summary>
    ''' Dạng hiển thị
    ''' </summary>
    Public CIPOutputOrder As String
    ''' <summary>
    ''' Độ dài mã
    ''' </summary>
    Public CIPOutputLength As Integer
    ''' <summary>
    ''' Dấu phân cách
    ''' </summary>
    Public CIPSeperated As Boolean
    ''' <summary>
    ''' S1Length
    ''' </summary>
    Public CIPS1Length As Integer
    ''' <summary>
    ''' S2Length
    ''' </summary>
    Public CIPS2Length As Integer
    ''' <summary>
    ''' S3Length
    ''' </summary>
    Public CIPS3Length As Integer
    ''' <summary>
    ''' Dấu phân cách
    ''' </summary>
    Public CIPSeperator As String
    ''' <summary>
    ''' Độ dài số
    ''' </summary>
    Public CIPAutoNumberLength As Integer
    ''' <summary>
    ''' Bắt buộc nhập người tiếp nhận
    ''' </summary>
    Public ObligatoryReceiver As Boolean
    ''' <summary>
    ''' Dung dự án hạng mục
    ''' </summary>
    Public UseD54ForCIP As Integer
    ''' <summary>
    ''' Dùng ngân sách hạng mục ngân sách
    ''' </summary>
    Public UseBudgetForCIP As Integer
    ''' <summary>
    ''' Tập hợp XDCB cho mã BĐS
    ''' </summary>
    Public CIPforPropertyProduct As Boolean
    '------------------------------------------------------------------------
    ''' <summary>
    ''' Cho phép bổ sung tài sản theo bộ định mức
    ''' </summary>
    Public IsCheckNormIDAllocated As Boolean
    Public IsShowFormAutoCreate As Boolean
    Public AllowChangeDivision As Boolean
    Public IsAllowChangeAccount As Boolean
    Public IsCalDepByDate As Boolean
    Public IsEditAnaID As Boolean
    '------------------------------------------------------------------------
    Public IsCalPeriodByRate As Boolean
    Public IsObligatoryManagement As Boolean

    '  D02 Systems here
    '------------------------------------------------------------------------
    Public LoadSystem As Boolean
End Structure

#End Region

Public Module D02X9940
    ''' <summary>
    ''' Lưu trữ các thiết lập tùy chọn
    ''' </summary>
    Public D02Options As StructureOption
    ''' <summary>
    ''' Lưu trữ các thiết lập Thông tin hệ thống
    ''' </summary>
    Public D02Systems As StructureSystem
    ''' <summary>
    ''' Lưu trữ các thiết lập format
    ''' </summary>
    Public Const D02 As String = "D02"
    Public Const MODULED02 As String = "D02D0040"
    Public gnDecreaseAsset As Byte
    Public giPerF5558 As Integer


    Public Sub LoadOptions()
        Try
          
            Dim clsOptions As New Lemon3.clsOptions(D02)
            'With D02Options
            '    'Kiểm tra tồn tại đường dẫn mới lưu .Net thì lấy dữ liệu, ngược lại thì lấy theo đường dẫn cũ (Lemon3_Dxx)
            '    'Kiem tra ky cac ten luu xuong cua VB6 de gan vao NET
            '    If D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "MessageAskBeforeSave") = "" Then 'Lay duong dan cu VB6
            '        Dim D02LocalOptionsLocations As String = "D02"
            '        Dim Options As String = "Options"

            '        With D02Options
            '            .DefaultDivisionID = GetSetting(D02LocalOptionsLocations, Options, "Division", "")
            '            .MessageAskBeforeSave = CType(GetSetting(D02LocalOptionsLocations, Options, "AskBeforeSave", "True"), Boolean)
            '            .MessageWhenSaveOK = CType(GetSetting(D02LocalOptionsLocations, Options, "MessageWhenSaveOK", "True"), Boolean)
            '            .SaveLastRecent = CType(GetSetting(D02LocalOptionsLocations, Options, "SaveRecentValues", "False"), Boolean)
            '            .RoundConvertedAmount = CType(GetSetting(D02LocalOptionsLocations, Options, "RoundConvertedAmount", "False"), Boolean)
            '            .LockConvertedAmount = CType(GetSetting(D02LocalOptionsLocations, Options, "LockConvertedAmount", "False"), Boolean)
            '            .ViewFormPeriodWhenAppRun = CType(GetSetting(D02LocalOptionsLocations, Options, "AcountingScreen", "False"), Boolean)
            '            .ReportLanguage = CType(GetSetting(D02LocalOptionsLocations, Options, "nRPLang", "0"), Byte)
            '            .ViewWorkflow = CType(GetSetting(D02LocalOptionsLocations, Options, "ShowDiagramTransaction", "False"), Boolean)
            '            .ShowReportPath = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ShowReportPath", "True"), Boolean)

            '        End With
            '    Else 'Lấy đường dẫn mới .Net
            With D02Options
                '.DefaultDivisionID = D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "DefaultDivisionID", "")
                '.MessageAskBeforeSave = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "MessageAskBeforeSave", "True"), Boolean)
                '.MessageWhenSaveOK = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "MessageWhenSaveOK", "True"), Boolean)
                '.SaveLastRecent = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "SaveLastRecent", "False"), Boolean)
                '.RoundConvertedAmount = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "RoundConvertedAmount", "False"), Boolean)
                '.LockConvertedAmount = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "LockConvertedAmount", "False"), Boolean)
                '.ViewFormPeriodWhenAppRun = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ViewFormPeriodWhenAppRun", "False"), Boolean)
                '.ReportLanguage = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ReportLanguage", "0"), Byte)
                '.ViewWorkflow = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ViewWorkflow", "False"), Boolean)
                '.ShowReportPath = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ShowReportPath", "True"), Boolean)

                .DefaultDivisionID = L3String(clsOptions.GetValue("DefaultDivisionID", ""))
                .MessageAskBeforeSave = L3Bool(clsOptions.GetValue("MessageAskBeforeSave", True))
                .MessageWhenSaveOK = L3Bool(clsOptions.GetValue("MessageWhenSaveOK", False))
                .SaveLastRecent = L3Bool(clsOptions.GetValue("SaveLastRecent", False))
                .RoundConvertedAmount = L3Bool(clsOptions.GetValue("RoundConvertedAmount", False))
                .LockConvertedAmount = L3Bool(clsOptions.GetValue("LockConvertedAmount", False))
                .ViewFormPeriodWhenAppRun = L3Bool(clsOptions.GetValue("ViewFormPeriodWhenAppRun", False))
                .ReportLanguage = L3Byte(clsOptions.GetValue("ReportLanguage", 0))
                .ViewWorkflow = L3Bool(clsOptions.GetValue("ViewWorkflow", False))
                .ShowReportPath = L3Bool(clsOptions.GetValue("ShowReportPath", True))
            End With
            '    End If
            'End With
        Catch ex As Exception
            D99C0008.Msg(ex.StackTrace)
        End Try
    End Sub

    ''' <summary>
    ''' Load toàn bộ các thống số thiết lập hệ thống vào biến D02Systems
    ''' </summary>
    Public Sub LoadSystems()

        Try
            Dim sSQL As String = "Select * From D02T0000"
            Dim dt As DataTable = ReturnDataTable(sSQL)
            If dt.Rows.Count > 0 Then
                With D02Systems
                    .DefaultDivisionID = dt.Rows(0).Item("DefaultDivisionID").ToString
                    ' TK tài sản
                    .DefAssetAccountID = dt.Rows(0).Item("DefAssetAccountID").ToString
                    ' TK khấu hao
                    .DefDepreciationAccountID = dt.Rows(0).Item("DefDepreciationAccountID").ToString
                    ' Nguồn vốn
                    .DefSourceID = dt.Rows(0).Item("DefSourceID").ToString
                    ' Phân bổ KH
                    .DefAssignmentID = dt.Rows(0).Item("DefAssignmentID").ToString
                    ' Phương pháp KH
                    .MethodID = dt.Rows(0).Item("MethodID").ToString
                    ' Xử lý KH kỳ cuối
                    .MethodEndID = dt.Rows(0).Item("MethodEndID").ToString
                    ' Các bút toán giảm TS
                    .DecreaseAsset = L3Bool(dt.Rows(0).Item("DecreaseAsset"))
                    ' Tạo mã tự động cho tài sản cố định
                    .AssetAuto = L3Int(dt.Rows(0).Item("AssetAuto"))
                    ' Phân loại 1
                    .AssetS1Enabled = L3Bool(dt.Rows(0).Item("AssetS1Enabled"))
                    ' Phân loại 2
                    .AssetS2Enabled = L3Bool(dt.Rows(0).Item("AssetS2Enabled"))
                    ' Phân loại 3
                    .AssetS3Enabled = L3Bool(dt.Rows(0).Item("AssetS3Enabled"))
                    ' AssetS1Default
                    .AssetS1Default = dt.Rows(0).Item("AssetS1Default").ToString
                    ' AssetS2Default
                    .AssetS2Default = dt.Rows(0).Item("AssetS2Default").ToString
                    ' AssetS3Default
                    .AssetS3Default = dt.Rows(0).Item("AssetS3Default").ToString
                    ' S1Length
                    .S1Length = L3Int(dt.Rows(0).Item("S1Length"))
                    ' S2Length
                    .S2Length = L3Int(dt.Rows(0).Item("S2Length"))
                    ' S3Length
                    .S3Length = L3Int(dt.Rows(0).Item("S3Length"))
                    ' Dấu phân cách
                    .AssetSeperated = L3Bool(dt.Rows(0).Item("AssetSeperated"))
                    ' Dấu phân cách
                    .AssetSeperator = dt.Rows(0).Item("AssetSeperator").ToString
                    ' Dạng hiển thị
                    .AssetOutputOrder = dt.Rows(0).Item("AssetOutputOrder").ToString
                    ' Độ dài số
                    .AutoNumberLength = L3Int(dt.Rows(0).Item("AutoNumberLength"))
                    ' Độ dài mã
                    .AssetOutputLength = L3Int(dt.Rows(0).Item("AssetOutputLength"))
                    'Tạo mã tự động Incident 69247
                    .CIPAuto = L3Int(dt.Rows(0).Item("CIPAuto"))
                    .IsAssetIDForD02D43 = L3Bool(dt.Rows(0).Item("IsAssetIDForD02D43"))
                    .UseProperty = L3Bool(dt.Rows(0).Item("UseProperty"))
                    'Phan danh cho XDCB
                    .CIPS1Enabled = L3Bool(dt.Rows(0).Item("CIPS1Enabled"))
                    .CIPS2Enabled = L3Bool(dt.Rows(0).Item("CIPS2Enabled"))
                    .CIPS3Enabled = L3Bool(dt.Rows(0).Item("CIPS3Enabled"))
                    .CIPS1Default = dt.Rows(0).Item("CIPS1Default").ToString
                    .CIPS2Default = dt.Rows(0).Item("CIPS2Default").ToString
                    .CIPS3Default = dt.Rows(0).Item("CIPS3Default").ToString
                    .CIPOutputOrder = dt.Rows(0).Item("CIPOutputOrder").ToString
                    .CIPOutputLength = L3Int(dt.Rows(0).Item("CIPOutputLength"))
                    .CIPSeperated = L3Bool(dt.Rows(0).Item("CIPSeperated"))
                    .CIPS1Length = L3Int(dt.Rows(0).Item("CIPS1Length"))
                    .CIPS2Length = L3Int(dt.Rows(0).Item("CIPS2Length"))
                    .CIPS3Length = L3Int(dt.Rows(0).Item("CIPS3Length"))
                    .CIPSeperator = dt.Rows(0).Item("CIPSeperator").ToString
                    .CIPAutoNumberLength = L3Int(dt.Rows(0).Item("CIPAutoNumberLength"))
                    .ObligatoryReceiver = L3Bool(dt.Rows(0).Item("ObligatoryReceiver"))
                    .UseD54ForCIP = L3Int(dt.Rows(0).Item("UseD54ForCIP"))
                    .UseBudgetForCIP = L3Int(dt.Rows(0).Item("UseBudgetForCIP"))
                    .CIPforPropertyProduct = L3Bool(dt.Rows(0).Item("CIPforPropertyProduct"))
                    .IsCheckNormIDAllocated = L3Bool(dt.Rows(0).Item("IsCheckNormIDAllocated"))
                    .IsShowFormAutoCreate = L3Bool(dt.Rows(0).Item("IsShowFormAutoCreate"))
                    .AllowChangeDivision = L3Bool(dt.Rows(0).Item("AllowChangeDivision"))
                    .IsAllowChangeAccount = L3Bool(dt.Rows(0).Item("IsAllowChangeAccount"))
                    .IsCalDepByDate = L3Bool(dt.Rows(0).Item("IsCalDepByDate")) '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
                    .IsEditAnaID = L3Bool(dt.Rows(0).Item("IsEditAnaID")) '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
                    .IsCalPeriodByRate = L3Bool(dt.Rows(0).Item("IsCalPeriodByRate")) '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
                    .IsObligatoryManagement = L3Bool(dt.Rows(0).Item("IsObligatoryManagement")) '15/11/2022 - ID 250796 : Bổ sung checkbox bắt buộc nhập Bộ phận quản lý
                    .LoadSystem = True
                End With
                gnDecreaseAsset = CByte(dt.Rows(0).Item("DecreaseAsset").ToString)
                dt.Dispose()
            Else
                With D02Systems
                    .DefaultDivisionID = ""
                    ' TK tài sản
                    .DefAssetAccountID = ""
                    ' TK khấu hao
                    .DefDepreciationAccountID = ""
                    ' Nguồn vốn
                    .DefSourceID = ""
                    ' Phân bổ KH
                    .DefAssignmentID = ""
                    ' Phương pháp KH
                    .MethodID = ""
                    ' Xử lý KH kỳ cuối
                    .MethodEndID = ""
                    ' Các bút toán giảm TS
                    .DecreaseAsset = False
                    ' Tạo mã tự động cho tài sản cố định
                    .AssetAuto = 0
                    ' Phân loại 1
                    .AssetS1Enabled = False
                    ' Phân loại 2
                    .AssetS2Enabled = False
                    ' Phân loại 3
                    .AssetS3Enabled = False
                    ' AssetS1Default
                    .AssetS1Default = ""
                    ' AssetS2Default
                    .AssetS2Default = ""
                    ' AssetS3Default
                    .AssetS3Default = ""
                    ' S1Length
                    .S1Length = 0
                    ' S2Length
                    .S2Length = 0
                    ' S3Length
                    .S3Length = 0
                    ' Dấu phân cách
                    .AssetSeperated = False
                    ' Dấu phân cách
                    .AssetSeperator = ""
                    ' Dạng hiển thị
                    .AssetOutputOrder = ""
                    ' Độ dài số
                    .AutoNumberLength = 0
                    ' Độ dài mã
                    .AssetOutputLength = 0
                    .CIPAuto = 0
                    .IsAssetIDForD02D43 = False
                    .UseProperty = False
                    'Phan danh cho XDCB
                    .CIPS1Enabled = False
                    .CIPS2Enabled = False
                    .CIPS3Enabled = False
                    .CIPS1Default = ""
                    .CIPS2Default = ""
                    .CIPS3Default = ""
                    .CIPOutputOrder = ""
                    .CIPOutputLength = 0
                    .CIPSeperated = False
                    .CIPS1Length = 0
                    .CIPS2Length = 0
                    .CIPS3Length = 0
                    .CIPSeperator = ""
                    .CIPAutoNumberLength = 0
                    .ObligatoryReceiver = False
                    .UseD54ForCIP = 0
                    .UseBudgetForCIP = 0
                    .CIPforPropertyProduct = False
                    .IsCheckNormIDAllocated = False
                    .IsShowFormAutoCreate = False
                    .AllowChangeDivision = False
                    .IsAllowChangeAccount = False
                    .IsCalDepByDate = False '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
                    .IsEditAnaID = False '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
                    .IsCalPeriodByRate = False '31/3/2022, Bùi Thị Thanh Tuyền:id 214947-ORG - Phát triển khi hình thành tài sản cố định nhập tỷ lệ khấu hao (theo năm) thì tính ngược lại số kỳ, giá trị phân bổ
                    .IsObligatoryManagement = False '15/11/2022 - ID 250796 : Bổ sung checkbox bắt buộc nhập Bộ phận quản lý
                    .LoadSystem = False
                End With
                gnDecreaseAsset = 0
            End If
        Catch ex As Exception
            D99C0008.Msg(ex.StackTrace & vbCrLf & ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Hỏi trước khi lưu tùy thuộc vào thiết lập ở phần Tùy chọn
    ''' </summary>
    Public Function AskSave() As DialogResult
        If D02Options.MessageAskBeforeSave Then
            Return D99C0008.MsgAskSave()
        Else
            Return DialogResult.Yes
        End If
    End Function

    ''' <summary>
    ''' Thông báo trước khi khóa phiếu
    ''' </summary>    
    Public Function AskLocked() As DialogResult
        If D02Options.MessageAskBeforeSave Then
            Return D99C0008.MsgAsk(rL3("MSG000002"), MessageBoxDefaultButton.Button2)
        Else
            Return DialogResult.Yes
        End If
    End Function

    ''' <summary>
    ''' Thông báo cột đã bị khóa khi nhấn phím nóng trên cột này để copy, xóa
    ''' </summary>
    Public Function MsgLockedColumn() As String
        Dim sMsg As String = ""
        sMsg = rl3("Cot_nay_da_bi_khoa_khong_duoc_phep_thao_tac_tren_cot_nay") 'rl3("Cot_nay_da_bi_khoa_khong_duoc_phep_thao_tac_tren_cot_nay")
        Return sMsg

    End Function

    ''' <summary>
    ''' Thông báo khi lưu thành công tùy theo phần thiết lập ở tùy chọn
    ''' </summary>
    Public Sub SaveOK()
        If D02Options.MessageWhenSaveOK Then D99C0008.MsgSaveOK()
    End Sub

    ''' <summary>
    ''' Thông báo không khóa được phiếu
    ''' </summary>
    Public Sub LockedNotOK()
        D99C0008.MsgL3(rl3("MSG000003"))
    End Sub


    Public Sub LoadInfoGeneral()
        If D02Systems.LoadSystem Then Exit Sub
        GetCodeTable()
        GeneralItems()
        LoadUserKey()
        LoadCreateBy()
        giPerF5558 = ReturnPermission("D02F5558")
        LoadOptions()
        LoadSystems()
        LoadCustomFormat()
    End Sub

    Public Function ComboValue(ByVal c1Combo As C1.Win.C1List.C1Combo) As String
        If c1Combo.Text = "" Then Return ""
        If c1Combo.SelectedValue IsNot Nothing Then
            Return c1Combo.SelectedValue.ToString
        Else
            Return ""
        End If
    End Function

    Public Function GetReportPath(ByVal ReportTypeID As String, ByVal ReportName As String, ByVal CustomReport As String, ByRef ReportPath As String, Optional ByRef ReportTitle As String = "", Optional ByVal ModuleID As String = "02", Optional ByVal bPrintUnicode As Boolean = False) As String
        Return Lemon3.Reports.GetReportPath(ReportTypeID, ReportName, CustomReport, ReportPath, ReportTitle, ModuleID, D02Options.ShowReportPath, bPrintUnicode)
    End Function
End Module