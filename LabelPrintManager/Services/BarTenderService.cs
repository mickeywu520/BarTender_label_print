#define BARTENDER_SDK_AVAILABLE

using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ComponentModel;

#if BARTENDER_SDK_AVAILABLE
using Seagull.BarTender.Print;
#endif

namespace LabelPrintManager.Services
{
    /// <summary>
    /// BarTender SDK 服務類別
    /// </summary>
    public class BarTenderService
    {
#if BARTENDER_SDK_AVAILABLE
        // BarTender SDK 物件
        private Engine btEngine;
        private LabelFormatDocument btFormat;
#endif

        private string _currentBtwFilePath;
        private const string PRINTER_NAME = "lc01"; // 測試印表機名稱
        private const string PRINTER_IP = "192.168.0.240"; // 測試印表機IP
        private Dictionary<string, string> _fieldValues;
        private bool _sdkAvailable = false;
        private string _selectedPrinterName; // 選擇的印表機名稱

        // BackgroundWorker 相關
        public class PrintJobData
        {
            public string PrinterName { get; set; }
            public int Copies { get; set; }
        }

        public class PrintJobResult
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public string PrinterLocation { get; set; }
            public int Copies { get; set; }
            public Exception Error { get; set; }
        }

        public BarTenderService()
        {
            _fieldValues = new Dictionary<string, string>();

            // 清理舊的預覽暫存檔案
            CleanupOldPreviewFiles();

#if BARTENDER_SDK_AVAILABLE
            try
            {
                // 初始化 BarTender 引擎
                btEngine = new Engine(true);
                _sdkAvailable = true;
                Console.WriteLine("BarTender SDK 初始化成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 BarTender 引擎失敗: {ex.Message}\n將使用模擬模式", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _sdkAvailable = false;
            }
#else
            Console.WriteLine("BarTender SDK 不可用，使用模擬模式");
            _sdkAvailable = false;
#endif
        }

        /// <summary>
        /// 載入 BTW 檔案
        /// </summary>
        /// <param name="btwFilePath">BTW 檔案路徑</param>
        /// <returns>是否載入成功</returns>
        public bool LoadBtwFile(string btwFilePath)
        {
            try
            {
                if (!File.Exists(btwFilePath))
                {
                    MessageBox.Show($"找不到檔案：{btwFilePath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

#if BARTENDER_SDK_AVAILABLE
                // 關閉之前的格式
                if (btFormat != null)
                {
                    btFormat.Close(SaveOptions.DoNotSaveChanges);
                    btFormat = null;
                }
#endif

                _currentBtwFilePath = btwFilePath;

#if BARTENDER_SDK_AVAILABLE
                // 使用 BarTender SDK 載入 BTW 檔案
                if (_sdkAvailable && btEngine != null)
                {
                    btFormat = btEngine.Documents.Open(btwFilePath);

                    // 診斷：列出 BTW 檔案中的所有可用欄位
                    DiagnoseBtwFileStructure();

                    MessageBox.Show($"已載入 BTW 檔案：{Path.GetFileName(btwFilePath)}\n\n請查看控制台輸出以了解檔案結構", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"BarTender 引擎未初始化，使用模擬模式載入：{Path.GetFileName(btwFilePath)}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
#else
                MessageBox.Show($"已載入 BTW 檔案（模擬模式）：{Path.GetFileName(btwFilePath)}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入 BTW 檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 獲取標籤預覽圖片
        /// </summary>
        /// <returns>預覽圖片</returns>
        public Image GetLabelPreview()
        {
            try
            {
                if (string.IsNullOrEmpty(_currentBtwFilePath))
                    return null;

#if BARTENDER_SDK_AVAILABLE
                // 如果有真正的 BarTender 格式，使用 SDK 匯出
                if (_sdkAvailable && btFormat != null)
                {
                    return ExportLabelToImage();
                }
                else
                {
                    // 如果沒有 SDK 或載入失敗，使用模擬預覽
                    return GenerateMockLabelPreview();
                }
#else
                // 沒有 SDK，使用模擬預覽
                return GenerateMockLabelPreview();
#endif
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成預覽圖片時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // 如果 SDK 匯出失敗，嘗試使用模擬預覽
                try
                {
                    return GenerateMockLabelPreview();
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 使用 BarTender SDK 匯出標籤圖片
        /// </summary>
        /// <returns>標籤圖片</returns>
        private Image ExportLabelToImage()
        {
#if BARTENDER_SDK_AVAILABLE
            try
            {
                Console.WriteLine("使用 BarTender SDK 生成真正的標籤預覽...");

                // 使用 exe 執行檔當前目錄作為預覽圖片暫存路徑
                string exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string tempDir = Path.Combine(exeDirectory, "PreviewTemp", $"bt_preview_{DateTime.Now:yyyyMMdd_HHmmss}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    Console.WriteLine("使用 ExportPrintPreviewToFile 方法...");

                    // 使用官方範例的方法
                    Messages messages;
                    var result = btFormat.ExportPrintPreviewToFile(
                        tempDir,                                    // 輸出目錄
                        "Preview%PageNumber%.jpg",                  // 檔案名稱模式
                        ImageType.JPEG,                            // 圖片格式
                        Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, // 色彩深度
                        new Resolution(800, 600),                  // 解析度 (寬x高)
                        System.Drawing.Color.White,               // 背景色
                        OverwriteOptions.Overwrite,               // 覆寫選項
                        true,                                      // 包含背景
                        true,                                      // 包含邊框
                        out messages                               // 輸出訊息
                    );

                    Console.WriteLine($"ExportPrintPreviewToFile 結果: {result}");

                    // 檢查是否有訊息
                    if (messages != null && messages.Count > 0)
                    {
                        Console.WriteLine($"BarTender 訊息數量: {messages.Count}");
                        foreach (Seagull.BarTender.Print.Message message in messages)
                        {
                            Console.WriteLine($"  - {message.Text}");
                        }
                    }

                    // 檢查生成的檔案
                    string[] files = Directory.GetFiles(tempDir, "*.jpg");
                    Console.WriteLine($"生成的預覽檔案數量: {files.Length}");

                    if (files.Length > 0)
                    {
                        // 載入第一個預覽檔案
                        string previewFile = files[0];
                        Console.WriteLine($"載入預覽檔案: {Path.GetFileName(previewFile)}");

                        using (var fileStream = new FileStream(previewFile, FileMode.Open, FileAccess.Read))
                        {
                            var image = new Bitmap(fileStream);
                            Console.WriteLine($"預覽圖片尺寸: {image.Width}x{image.Height}");
                            return new Bitmap(image); // 創建副本以避免檔案鎖定
                        }
                    }
                    else
                    {
                        Console.WriteLine("沒有生成預覽檔案");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ExportPrintPreviewToFile 失敗: {ex.Message}");
                    Console.WriteLine($"錯誤詳細資訊: {ex}");
                }
                finally
                {
                    // 清理臨時目錄
                    try
                    {
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                            Console.WriteLine("已清理臨時目錄");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"清理臨時目錄失敗: {ex.Message}");
                    }
                }

                // 如果 SDK 方法失敗，回退到模擬預覽
                Console.WriteLine("BarTender SDK 匯出失敗，使用模擬預覽");
                return GenerateMockLabelPreview();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BarTender SDK 匯出過程發生錯誤: {ex.Message}");
                return GenerateMockLabelPreview();
            }
#else
            // 如果沒有 SDK，返回模擬圖片
            return GenerateMockLabelPreview();
#endif
        }

        /// <summary>
        /// 生成模擬的標籤預覽圖片
        /// </summary>
        /// <returns>模擬的標籤圖片</returns>
        private Image GenerateMockLabelPreview()
        {
            // 標籤實際尺寸 97.5 x 60 mm，按比例創建高解析度圖片
            // 使用 300 DPI，97.5mm = 3.84", 60mm = 2.36"
            // 3.84" * 300 DPI = 1152 pixels, 2.36" * 300 DPI = 708 pixels
            Bitmap bitmap = new Bitmap(1152, 708);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // 設定背景為白色
                g.Clear(Color.White);

                // 設定字體 (放大以適應高解析度)
                Font titleFont = new Font("Arial", 16, FontStyle.Bold);
                Font normalFont = new Font("Arial", 14);
                Font barcodeFont = new Font("Courier New", 12);

                Brush blackBrush = Brushes.Black;
                Pen blackPen = new Pen(Color.Black, 3);

                // 繪製外框
                g.DrawRectangle(blackPen, 15, 15, 1122, 678);

                // 繪製內部格線
                int y = 70;
                int rowHeight = 70;

                // 第一行：供應商和月份
                g.DrawRectangle(blackPen, 30, y, 450, rowHeight);
                g.DrawRectangle(blackPen, 480, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 780, y, 357, rowHeight);

                g.DrawString("供應商", titleFont, blackBrush, 35, y + 8);
                g.DrawString(GetFieldValue("SupplierName", "XX科技(股)公司"), normalFont, blackBrush, 35, y + 35);

                g.DrawString("月份", titleFont, blackBrush, 485, y + 8);
                g.DrawString(GetFieldValue("Months", "12"), titleFont, blackBrush, 485, y + 35);

                g.DrawString("ARBOR", titleFont, blackBrush, 785, y + 35);

                y += rowHeight;

                // 第二行：品名
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 987, rowHeight);

                g.DrawString("品名", titleFont, blackBrush, 35, y + 8);
                g.DrawString(GetFieldValue("ProductName", "FPC-7602 BOTTOM COVER BRACKET"), normalFont, blackBrush, 155, y + 35);

                y += rowHeight;

                // 第三行：單別和項次
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 450, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 750, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 870, y, 267, rowHeight);

                g.DrawString("單別", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("DocType", "3113"), normalFont, blackBrush, 155, y + 45);

                g.DrawString(GetFieldValue("DocType", "3113"), normalFont, blackBrush, 455, y + 35);

                g.DrawString("項次", titleFont, blackBrush, 755, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 875, y + 20);
                g.DrawString(GetFieldValue("DocItem", "0001"), normalFont, blackBrush, 875, y + 45);

                y += rowHeight;

                // 第四行：單號
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 987, rowHeight);

                g.DrawString("單號", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("DocNumber", "1140421033"), normalFont, blackBrush, 155, y + 45);

                y += rowHeight;

                // 第五行：料號
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 600, rowHeight);
                g.DrawRectangle(blackPen, 750, y, 387, rowHeight);

                g.DrawString("料號", titleFont, blackBrush, 35, y + 8);
                g.DrawString(GetFieldValue("ProductNumber", "371760236010SP"), normalFont, blackBrush, 155, y + 35);
                g.DrawString("||||||||||||||||||||||||||||||||||||", barcodeFont, blackBrush, 755, y + 20);

                y += rowHeight;

                // 第六行：數量
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 450, y, 567, rowHeight);
                g.DrawRectangle(blackPen, 1017, y, 120, rowHeight);

                g.DrawString("數量", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("Quantity", "1000"), normalFont, blackBrush, 455, y + 35);
                g.DrawString("PCS", normalFont, blackBrush, 1022, y + 35);

                y += rowHeight;

                // 第七行：D/C 和備註
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 450, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 750, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 870, y, 267, rowHeight);

                g.DrawString("D/C", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("DC", "250601"), normalFont, blackBrush, 455, y + 35);
                g.DrawString("備註", titleFont, blackBrush, 755, y + 8);

                y += rowHeight;

                // 第八行：HW Ver 和 QR Code
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 450, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 750, y, 387, rowHeight);

                g.DrawString("HW Ver", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("HWVer", "V1.01"), normalFont, blackBrush, 455, y + 35);

                // 繪製 QR Code 區域
                g.FillRectangle(Brushes.Black, 780, y + 15, 60, 60);
                g.FillRectangle(Brushes.White, 790, y + 25, 40, 40);
                g.DrawString("QR", normalFont, blackBrush, 800, y + 40);

                y += rowHeight;

                // 第九行：FW Ver
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 300, rowHeight);
                g.DrawRectangle(blackPen, 450, y, 687, rowHeight);

                g.DrawString("FW Ver", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("FWVer", "V1.6F"), normalFont, blackBrush, 455, y + 35);

                y += rowHeight;

                // 第十行：日期
                g.DrawRectangle(blackPen, 30, y, 120, rowHeight);
                g.DrawRectangle(blackPen, 150, y, 987, rowHeight);

                g.DrawString("日期", titleFont, blackBrush, 35, y + 8);
                g.DrawString("||||||||||||||||||||", barcodeFont, blackBrush, 155, y + 20);
                g.DrawString(GetFieldValue("Date", "2025/06/06"), normalFont, blackBrush, 155, y + 45);

                // 清理資源
                titleFont.Dispose();
                normalFont.Dispose();
                barcodeFont.Dispose();
                blackPen.Dispose();
            }

            return bitmap;
        }

        /// <summary>
        /// 獲取欄位值
        /// </summary>
        private string GetFieldValue(string fieldName, string defaultValue)
        {
            return _fieldValues.ContainsKey(fieldName) ? _fieldValues[fieldName] : defaultValue;
        }

        /// <summary>
        /// 獲取標籤預覽資訊（文字版本）
        /// </summary>
        /// <returns>預覽資訊字串</returns>
        public string GetLabelPreviewText()
        {
            if (string.IsNullOrEmpty(_currentBtwFilePath))
                return "未載入標籤檔案";

            string fileName = Path.GetFileName(_currentBtwFilePath);
            return $"標籤檔案: {fileName}\n\n" +
                   "預覽欄位:\n" +
                   "• 進貨單別\n" +
                   "• 進貨單號\n" +
                   "• 品號\n" +
                   "• 品名\n" +
                   "• 數量\n" +
                   "• D/C\n" +
                   "• HW Ver\n" +
                   "• FW Ver";
        }

        /// <summary>
        /// 設定標籤欄位值
        /// </summary>
        /// <param name="fieldName">欄位名稱</param>
        /// <param name="value">欄位值</param>
        public void SetFieldValue(string fieldName, string value)
        {
            try
            {
                // 儲存到字典中（用於模擬模式）
                _fieldValues[fieldName] = value;

#if BARTENDER_SDK_AVAILABLE
                // 如果有真正的 BarTender 格式，設定欄位值
                if (_sdkAvailable && btFormat != null)
                {
                    try
                    {
                        // 嘗試設定 SubStrings
                        bool fieldSet = TrySetSubString(fieldName, value);

                        if (!fieldSet)
                        {
                            // 如果 SubStrings 失敗，嘗試 DatabaseFields
                            fieldSet = TrySetDatabaseField(fieldName, value);
                        }

                        if (!fieldSet)
                        {
                            Console.WriteLine($"警告：在 BarTender 格式中找不到欄位 '{fieldName}'");
                            SuggestSimilarFieldNames(fieldName);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"設定 BarTender 欄位 {fieldName} 時發生錯誤: {ex.Message}");
                    }
                }
#endif

                Console.WriteLine($"設定欄位 {fieldName} = {value}");
            }
            catch (Exception ex)
            {
                throw new Exception($"設定欄位 {fieldName} 失敗: {ex.Message}", ex);
            }
        }



        /// <summary>
        /// 關閉當前格式檔案
        /// </summary>
        public void CloseFormat()
        {
            try
            {
#if BARTENDER_SDK_AVAILABLE
                if (btFormat != null)
                {
                    btFormat.Close(SaveOptions.DoNotSaveChanges);
                    btFormat = null;
                }
#endif

                _currentBtwFilePath = null;
                _fieldValues.Clear();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"關閉格式檔案時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// 嘗試設定 SubString 欄位
        /// </summary>
        private bool TrySetSubString(string fieldName, string value)
        {
#if BARTENDER_SDK_AVAILABLE
            try
            {
                var namedSubStrings = btFormat.SubStrings;
                if (namedSubStrings[fieldName] != null)
                {
                    namedSubStrings[fieldName].Value = value;
                    Console.WriteLine($"成功設定 SubString 欄位 '{fieldName}' = '{value}'");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"設定 SubString 欄位 '{fieldName}' 失敗: {ex.Message}");
            }
#endif
            return false;
        }

        /// <summary>
        /// 嘗試設定 DatabaseField 欄位
        /// </summary>
        private bool TrySetDatabaseField(string fieldName, string value)
        {
#if BARTENDER_SDK_AVAILABLE
            try
            {
                // BarTender 2019 R10 可能沒有 DatabaseFields 屬性
                // 暫時跳過這個方法
                Console.WriteLine($"DatabaseField 設定暫時不支援 (BarTender 2019 R10)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"設定 DatabaseField 欄位 '{fieldName}' 失敗: {ex.Message}");
            }
#endif
            return false;
        }

        /// <summary>
        /// 建議相似的欄位名稱
        /// </summary>
        private void SuggestSimilarFieldNames(string targetFieldName)
        {
#if BARTENDER_SDK_AVAILABLE
            try
            {
                Console.WriteLine($"正在尋找與 '{targetFieldName}' 相似的欄位名稱...");

                // 檢查 SubStrings
                var subStrings = btFormat.SubStrings;
                for (int i = 0; i < subStrings.Count; i++)
                {
                    try
                    {
                        string fieldName = subStrings[i].Name;
                        if (IsFieldNameSimilar(targetFieldName, fieldName))
                        {
                            Console.WriteLine($"  建議 SubString: '{fieldName}' (可能匹配 '{targetFieldName}')");
                        }
                    }
                    catch { }
                }

                // DatabaseFields 在 BarTender 2019 R10 中可能不可用
                Console.WriteLine("  DatabaseFields 檢查跳過 (BarTender 2019 R10 限制)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"建議欄位名稱時發生錯誤: {ex.Message}");
            }
#endif
        }

        /// <summary>
        /// 檢查兩個欄位名稱是否相似
        /// </summary>
        private bool IsFieldNameSimilar(string target, string candidate)
        {
            if (string.IsNullOrEmpty(target) || string.IsNullOrEmpty(candidate))
                return false;

            target = target.ToLower();
            candidate = candidate.ToLower();

            // 完全匹配
            if (target == candidate) return true;

            // 包含關係
            if (target.Contains(candidate) || candidate.Contains(target)) return true;

            // 檢查常見的對應關係
            var mappings = new Dictionary<string, string[]>
            {
                { "suppliername", new[] { "supplier", "廠商", "供應商" } },
                { "productname", new[] { "product", "品名", "產品" } },
                { "doctype", new[] { "type", "單別", "類型" } },
                { "docitem", new[] { "item", "項次", "項目" } },
                { "docnumber", new[] { "number", "單號", "編號" } },
                { "productcode", new[] { "code", "料號", "代碼" } },
                { "quantity", new[] { "qty", "數量", "amount" } },
                { "hwver", new[] { "hardware", "硬體", "hw" } },
                { "fwver", new[] { "firmware", "韌體", "fw" } },
                { "purchasedate", new[] { "date", "日期", "purchase" } }
            };

            if (mappings.ContainsKey(target))
            {
                foreach (string keyword in mappings[target])
                {
                    if (candidate.Contains(keyword)) return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 檢查欄位對應關係
        /// </summary>
        private void CheckFieldMapping(dynamic subStrings)
        {
#if BARTENDER_SDK_AVAILABLE
            // 程式中使用的欄位名稱
            var expectedFields = new Dictionary<string, string>
            {
                { "SupplierName", "供應商" },
                { "Months", "月份" },
                { "ProductName", "品名" },
                { "DocType", "單別" },
                { "DocItem", "項次" },
                { "DocNumber", "單號" },
                { "ProductNumber", "料號" },
                { "Quantity", "數量" },
                { "DC", "D/C" },
                { "HWVer", "HW Ver" },
                { "FWVer", "FW Ver" },
                { "Date", "日期" }
            };

            Console.WriteLine("檢查程式欄位與 BTW 檔案的對應關係：");

            foreach (var expectedField in expectedFields)
            {
                string fieldName = expectedField.Key;
                string description = expectedField.Value;

                try
                {
                    var subString = subStrings[fieldName];
                    if (subString != null)
                    {
                        Console.WriteLine($"  ✓ {description} ({fieldName}) - 找到，當前值: '{subString.Value}'");
                    }
                    else
                    {
                        Console.WriteLine($"  ✗ {description} ({fieldName}) - 未找到");
                    }
                }
                catch
                {
                    Console.WriteLine($"  ✗ {description} ({fieldName}) - 未找到");
                }
            }

            Console.WriteLine();
            Console.WriteLine("=== 對應關係檢查完成 ===");
#endif
        }

        /// <summary>
        /// 診斷 BTW 檔案結構 - 列出所有可用的欄位和物件
        /// </summary>
        private void DiagnoseBtwFileStructure()
        {
#if BARTENDER_SDK_AVAILABLE
            if (btFormat == null)
            {
                Console.WriteLine("=== BTW 檔案診斷 ===");
                Console.WriteLine("錯誤：btFormat 為 null，無法進行診斷");
                return;
            }

            try
            {
                Console.WriteLine("=== BTW 檔案結構診斷開始 ===");
                Console.WriteLine($"檔案路徑: {_currentBtwFilePath}");
                Console.WriteLine();

                // 1. 檢查 SubStrings (命名子字串)
                Console.WriteLine("=== 1. SubStrings (命名子字串) ===");
                try
                {
                    var subStrings = btFormat.SubStrings;
                    Console.WriteLine($"SubStrings 總數: {subStrings.Count}");

                    if (subStrings.Count > 0)
                    {
                        Console.WriteLine("找到的 SubStrings：");
                        for (int i = 0; i < subStrings.Count; i++)
                        {
                            try
                            {
                                var subString = subStrings[i];
                                Console.WriteLine($"  [{i}] 名稱: '{subString.Name}' | 當前值: '{subString.Value}'");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  [{i}] 讀取失敗: {ex.Message}");
                            }
                        }

                        Console.WriteLine();
                        Console.WriteLine("=== 欄位對應檢查 ===");
                        CheckFieldMapping(subStrings);
                    }
                    else
                    {
                        Console.WriteLine("  沒有找到 SubStrings");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"讀取 SubStrings 時發生錯誤: {ex.Message}");
                }

                Console.WriteLine();

                // 2. 檢查其他可能的屬性
                Console.WriteLine("=== 2. 其他屬性檢查 ===");
                try
                {
                    // 嘗試檢查 btFormat 的其他屬性
                    Console.WriteLine($"格式名稱: {btFormat.FileName}");
                    Console.WriteLine($"格式類型: {btFormat.GetType().Name}");

                    // 列出 btFormat 的所有公開屬性
                    var properties = btFormat.GetType().GetProperties();
                    Console.WriteLine($"可用屬性總數: {properties.Length}");
                    foreach (var prop in properties)
                    {
                        try
                        {
                            Console.WriteLine($"  屬性: {prop.Name} (類型: {prop.PropertyType.Name})");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  屬性: {prop.Name} - 讀取失敗: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"檢查其他屬性時發生錯誤: {ex.Message}");
                }

                Console.WriteLine();
                Console.WriteLine("=== BTW 檔案結構診斷完成 ===");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"診斷過程中發生錯誤: {ex.Message}");
                Console.WriteLine($"錯誤詳細資訊: {ex}");
            }
#else
            Console.WriteLine("BarTender SDK 不可用，無法進行診斷");
#endif
        }

        /// <summary>
        /// 清理舊的預覽暫存檔案
        /// </summary>
        private void CleanupOldPreviewFiles()
        {
            try
            {
                string exeDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string previewTempDir = Path.Combine(exeDirectory, "PreviewTemp");

                if (Directory.Exists(previewTempDir))
                {
                    Console.WriteLine("清理舊的預覽暫存檔案...");

                    // 獲取所有預覽目錄
                    string[] previewDirs = Directory.GetDirectories(previewTempDir, "bt_preview_*");

                    foreach (string dir in previewDirs)
                    {
                        try
                        {
                            // 檢查目錄是否超過 1 小時
                            DirectoryInfo dirInfo = new DirectoryInfo(dir);
                            if (DateTime.Now - dirInfo.CreationTime > TimeSpan.FromHours(1))
                            {
                                Directory.Delete(dir, true);
                                Console.WriteLine($"已清理舊預覽目錄: {Path.GetFileName(dir)}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"清理預覽目錄 {dir} 失敗: {ex.Message}");
                        }
                    }

                    // 如果 PreviewTemp 目錄為空，也刪除它
                    if (Directory.GetDirectories(previewTempDir).Length == 0 &&
                        Directory.GetFiles(previewTempDir).Length == 0)
                    {
                        try
                        {
                            Directory.Delete(previewTempDir);
                            Console.WriteLine("已清理空的 PreviewTemp 目錄");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"清理 PreviewTemp 目錄失敗: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理預覽暫存檔案時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// 獲取可用印表機清單
        /// </summary>
        /// <returns>印表機名稱清單</returns>
        public List<string> GetAvailablePrinters()
        {
            List<string> printerNames = new List<string>();

#if BARTENDER_SDK_AVAILABLE
            try
            {
                if (_sdkAvailable)
                {
                    Printers printers = new Printers();
                    foreach (Printer printer in printers)
                    {
                        printerNames.Add(printer.PrinterName);
                        Console.WriteLine($"發現印表機: {printer.PrinterName}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取印表機清單時發生錯誤: {ex.Message}");
            }
#endif

            // 如果沒有找到印表機，添加預設測試印表機
            if (printerNames.Count == 0)
            {
                printerNames.Add(PRINTER_NAME);
                Console.WriteLine($"使用預設測試印表機: {PRINTER_NAME}");
            }

            return printerNames;
        }

        /// <summary>
        /// 獲取預設印表機名稱
        /// </summary>
        /// <returns>預設印表機名稱</returns>
        public string GetDefaultPrinter()
        {
#if BARTENDER_SDK_AVAILABLE
            try
            {
                if (_sdkAvailable)
                {
                    Printers printers = new Printers();
                    if (printers.Default != null)
                    {
                        Console.WriteLine($"預設印表機: {printers.Default.PrinterName}");
                        return printers.Default.PrinterName;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取預設印表機時發生錯誤: {ex.Message}");
            }
#endif

            // 如果無法獲取預設印表機，返回測試印表機
            return PRINTER_NAME;
        }

        /// <summary>
        /// 設定選擇的印表機
        /// </summary>
        /// <param name="printerName">印表機名稱</param>
        public void SetSelectedPrinter(string printerName)
        {
            _selectedPrinterName = printerName;
            Console.WriteLine($"設定印表機: {printerName}");
        }

        /// <summary>
        /// 獲取目前選擇的印表機
        /// </summary>
        /// <returns>印表機名稱</returns>
        public string GetSelectedPrinter()
        {
            return _selectedPrinterName ?? GetDefaultPrinter();
        }

        /// <summary>
        /// 獲取印表機的網路位置資訊
        /// </summary>
        /// <param name="printerName">印表機名稱</param>
        /// <returns>印表機位置</returns>
        private string GetPrinterLocation(string printerName)
        {
            try
            {
                // 如果是測試印表機，返回已知的IP位置
                if (printerName == PRINTER_NAME)
                {
                    return $"\\\\{PRINTER_IP}\\{PRINTER_NAME}";
                }

                // 對於其他印表機，嘗試推測網路位置
                if (printerName.Contains("192.168.0.240") || printerName.ToLower().Contains("lc01"))
                {
                    return $"\\\\{PRINTER_IP}\\{PRINTER_NAME}";
                }

                // 預設返回印表機名稱
                return printerName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取印表機位置時發生錯誤: {ex.Message}");
                return printerName;
            }
        }

        /// <summary>
        /// 背景執行列印工作（供 BackgroundWorker 使用）
        /// </summary>
        /// <param name="jobData">列印工作資料</param>
        /// <param name="worker">BackgroundWorker 實例</param>
        /// <returns>列印結果</returns>
        public PrintJobResult PrintLabelBackground(PrintJobData jobData, BackgroundWorker worker)
        {
            var result = new PrintJobResult
            {
                PrinterLocation = GetPrinterLocation(jobData.PrinterName),
                Copies = jobData.Copies
            };

            try
            {
                worker?.ReportProgress(10, "正在準備列印工作...");

#if BARTENDER_SDK_AVAILABLE
                if (!_sdkAvailable || btFormat == null)
                {
                    result.Success = false;
                    result.Message = "BarTender SDK 不可用或未載入格式檔案";
                    return result;
                }

                worker?.ReportProgress(30, "正在設定印表機...");

                // 設定印表機
                btFormat.PrintSetup.PrinterName = jobData.PrinterName;
                Console.WriteLine($"設定列印印表機: {jobData.PrinterName}");

                worker?.ReportProgress(50, "正在設定列印份數...");

                // 設定列印份數
                if (btFormat.PrintSetup.SupportsIdenticalCopies)
                {
                    btFormat.PrintSetup.IdenticalCopiesOfLabel = jobData.Copies;
                    Console.WriteLine($"設定列印份數: {jobData.Copies}");
                }

                worker?.ReportProgress(70, "正在發送列印工作...");

                // 執行列印 - 不等待完成，立即返回
                Console.WriteLine("發送列印工作到印表機...");
                Result printResult = btFormat.Print(jobData.PrinterName);

                worker?.ReportProgress(90, "列印工作已發送");

                Console.WriteLine("列印工作已發送到印表機");
                result.Success = true;
                result.Message = "列印工作已成功發送到印表機";

                worker?.ReportProgress(100, "完成");

                return result;
#else
                worker?.ReportProgress(50, "模擬列印模式...");
                System.Threading.Thread.Sleep(1000); // 模擬處理時間

                result.Success = true;
                result.Message = "模擬列印完成（SDK 未編譯）";
                worker?.ReportProgress(100, "完成");

                return result;
#endif
            }
            catch (Exception ex)
            {
                Console.WriteLine($"背景列印時發生錯誤: {ex.Message}");
                result.Success = false;
                result.Message = $"列印時發生錯誤: {ex.Message}";
                result.Error = ex;
                return result;
            }
        }

        /// <summary>
        /// 列印標籤（保留原有方法以向後相容）
        /// </summary>
        /// <param name="copies">列印份數</param>
        /// <returns>列印是否成功</returns>
        public bool PrintLabel(int copies = 1)
        {
#if BARTENDER_SDK_AVAILABLE
            try
            {
                if (!_sdkAvailable || btFormat == null)
                {
                    Console.WriteLine("BarTender SDK 不可用或未載入格式檔案");
                    return false;
                }

                // 設定印表機
                string printerName = GetSelectedPrinter();
                btFormat.PrintSetup.PrinterName = printerName;
                Console.WriteLine($"設定列印印表機: {printerName}");

                // 設定列印份數
                if (btFormat.PrintSetup.SupportsIdenticalCopies)
                {
                    btFormat.PrintSetup.IdenticalCopiesOfLabel = copies;
                    Console.WriteLine($"設定列印份數: {copies}");
                }

                // 執行列印 - 不等待完成，立即返回
                Console.WriteLine("發送列印工作到印表機...");

                // 使用不等待完成的列印方式
                Result result = btFormat.Print(printerName);

                // 立即顯示成功訊息，不等待印表機回應
                Console.WriteLine("列印工作已發送到印表機");

                // 獲取印表機的網路位置資訊
                string printerLocation = GetPrinterLocation(printerName);

                MessageBox.Show($"✅ 已傳送至印表機\n\n" +
                              $"印表機：{printerName}\n" +
                              $"位置：{printerLocation}\n" +
                              $"份數：{copies}\n\n" +
                              "列印工作已發送，請稍候取件。",
                              "列印已發送",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"列印標籤時發生錯誤: {ex.Message}");
                MessageBox.Show($"列印標籤時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
#else
            Console.WriteLine("BarTender SDK 不可用，無法列印");
            MessageBox.Show("BarTender SDK 不可用，無法列印", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
#endif
        }

        /// <summary>
        /// 釋放資源
        /// </summary>
        public void Dispose()
        {
            try
            {
                CloseFormat();

#if BARTENDER_SDK_AVAILABLE
                if (btEngine != null)
                {
                    btEngine.Stop();
                    btEngine = null;
                }
#endif

                // 最後清理一次預覽暫存檔案
                CleanupOldPreviewFiles();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"釋放 BarTender 資源時發生錯誤: {ex.Message}");
            }
        }
    }
}
