using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LabelPrintManager.Services;
using LabelPrintManager.Models;

namespace LabelPrintManager
{
    public partial class Form1 : Form
    {
        private ApiService _apiService;
        private BarTenderService _barTenderService;
        private PurchaseReceiptModel _currentReceiptData;

        // BackgroundWorker 相關
        private BackgroundWorker _printWorker;
        private ProgressForm _progressForm;
        private readonly object _previewLock = new object(); // 預覽更新鎖定

        public Form1()
        {
            InitializeComponent();
            InitializeServices();

            // 初始化控制項狀態（步驟式防呆）
            InitializeControlStates();
        }

        /// <summary>
        /// 初始化服務
        /// </summary>
        private void InitializeServices()
        {
            _apiService = new ApiService();
            _barTenderService = new BarTenderService();
        }



        /// <summary>
        /// 初始化控制項狀態（步驟式防呆機制）
        /// </summary>
        private void InitializeControlStates()
        {
            Console.WriteLine("=== 初始化控制項狀態 ===");

            // 步驟1：未載入 BTW 檔案時，禁用所有後續功能
            SetStep1State(); // 只有 BTW 檔案選擇可用

            Console.WriteLine("控制項狀態初始化完成");
        }

        /// <summary>
        /// 步驟1狀態：只有 BTW 檔案選擇可用
        /// </summary>
        private void SetStep1State()
        {
            Console.WriteLine("設定步驟1狀態：只有 BTW 檔案選擇可用");

            // BTW 檔案選擇 - 可用
            buttonBrowseBtw.Enabled = true;
            textBoxBtwPath.Enabled = true;

            // API 資料輸入區域 - 禁用
            textBoxDocType.Enabled = false;
            textBoxDocNumber.Enabled = false;
            textBoxDocItem.Enabled = false;
            comboBoxDocCategory.Enabled = false;
            buttonGetData.Enabled = false;

            // 列印設定區域 - 禁用
            textBoxQuantity.Enabled = false;
            textBoxDc.Enabled = false;
            textBoxHwVer.Enabled = false;
            textBoxFwVer.Enabled = false;
            textBoxProductName.Enabled = false;
            comboBoxPrinter.Enabled = false;
            textBoxCopies.Enabled = false;
            buttonUpdatePreview.Enabled = false;
            buttonPrint.Enabled = false;

            // 初始化印表機清單
            InitializePrinters();

            // 設定視覺提示
            SetGroupBoxState(groupBoxApiData, false, "請先載入 BTW 檔案");
            SetGroupBoxState(groupBoxPrintSettings, false, "請先載入 BTW 檔案並獲取資料");
        }

        /// <summary>
        /// 步驟2狀態：BTW 檔案已載入，開啟 API 資料輸入
        /// </summary>
        private void SetStep2State()
        {
            Console.WriteLine("設定步驟2狀態：BTW 檔案已載入，開啟 API 資料輸入");

            // BTW 檔案選擇 - 保持可用
            buttonBrowseBtw.Enabled = true;
            textBoxBtwPath.Enabled = true;

            // API 資料輸入區域 - 開啟
            textBoxDocType.Enabled = true;
            textBoxDocNumber.Enabled = true;
            textBoxDocItem.Enabled = true;
            comboBoxDocCategory.Enabled = true;
            buttonGetData.Enabled = true;

            // 列印設定區域 - 仍然禁用
            textBoxQuantity.Enabled = false;
            textBoxDc.Enabled = false;
            textBoxHwVer.Enabled = false;
            textBoxFwVer.Enabled = false;
            textBoxProductName.Enabled = false;
            comboBoxPrinter.Enabled = false;
            textBoxCopies.Enabled = false;
            buttonUpdatePreview.Enabled = false;
            buttonPrint.Enabled = false;

            // 設定視覺提示
            SetGroupBoxState(groupBoxApiData, true, "API 資料輸入");
            SetGroupBoxState(groupBoxPrintSettings, false, "請先獲取 API 資料");
        }

        /// <summary>
        /// 步驟3狀態：API 資料已獲取，開啟列印設定
        /// </summary>
        private void SetStep3State()
        {
            Console.WriteLine("設定步驟3狀態：API 資料已獲取，開啟列印設定");

            // BTW 檔案選擇 - 保持可用
            buttonBrowseBtw.Enabled = true;
            textBoxBtwPath.Enabled = true;

            // API 資料輸入區域 - 保持可用
            textBoxDocType.Enabled = true;
            textBoxDocNumber.Enabled = true;
            textBoxDocItem.Enabled = true;
            comboBoxDocCategory.Enabled = true;
            buttonGetData.Enabled = true;

            // 列印設定區域 - 全部開啟
            textBoxQuantity.Enabled = true;
            textBoxDc.Enabled = true;
            textBoxHwVer.Enabled = true;
            textBoxFwVer.Enabled = true;
            textBoxProductName.Enabled = true;
            comboBoxPrinter.Enabled = true;
            textBoxCopies.Enabled = true;
            buttonUpdatePreview.Enabled = true;
            buttonPrint.Enabled = true;

            // 設定視覺提示
            SetGroupBoxState(groupBoxApiData, true, "API 資料輸入");
            SetGroupBoxState(groupBoxPrintSettings, true, "列印設定");
        }

        /// <summary>
        /// 設定 GroupBox 的狀態和標題
        /// </summary>
        private void SetGroupBoxState(GroupBox groupBox, bool enabled, string title)
        {
            if (groupBox != null)
            {
                groupBox.Text = title;
                groupBox.ForeColor = enabled ? SystemColors.ControlText : SystemColors.GrayText;
            }
        }

        /// <summary>
        /// BTW 檔案瀏覽按鈕點擊事件
        /// </summary>
        private void buttonBrowseBtw_Click(object sender, EventArgs e)
        {
            if (openFileDialogBtw.ShowDialog() == DialogResult.OK)
            {
                string selectedFile = openFileDialogBtw.FileName;
                textBoxBtwPath.Text = selectedFile;

                // 載入 BTW 檔案
                if (_barTenderService.LoadBtwFile(selectedFile))
                {
                    // 顯示標籤預覽圖片
                    ShowLabelPreview();

                    // 步驟2：BTW 檔案載入成功，開啟 API 資料輸入
                    SetStep2State();

                    Console.WriteLine("BTW 檔案載入成功，已開啟 API 資料輸入功能");
                }
                else
                {
                    pictureBoxPreview.Image = null;
                    pictureBoxPreview.BackColor = Color.LightCoral;

                    // 載入失敗，回到步驟1
                    SetStep1State();

                    Console.WriteLine("BTW 檔案載入失敗，回到初始狀態");
                }
            }
        }

        /// <summary>
        /// 獲取資料按鈕點擊事件
        /// </summary>
        private async void buttonGetData_Click(object sender, EventArgs e)
        {
            try
            {
                // 驗證輸入
                if (string.IsNullOrWhiteSpace(textBoxDocType.Text) ||
                    string.IsNullOrWhiteSpace(textBoxDocNumber.Text) ||
                    string.IsNullOrWhiteSpace(textBoxDocItem.Text))
                {
                    MessageBox.Show("請填入完整的進貨單資訊", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 禁用按鈕防止重複點擊
                buttonGetData.Enabled = false;
                buttonGetData.Text = "獲取中...";

                // 判斷是否為托外進貨
                bool isSubcontract = comboBoxDocCategory.SelectedIndex == 1; // 0=一般進貨, 1=托外進貨

                // 呼叫 API 獲取資料
                _currentReceiptData = await _apiService.GetPurchaseReceiptAsync(
                    textBoxDocType.Text.Trim(),
                    textBoxDocNumber.Text.Trim(),
                    textBoxDocItem.Text.Trim(),
                    isSubcontract
                );

                if (_currentReceiptData != null)
                {
                    // 顯示 API 回傳資料
                    DisplayApiResult(_currentReceiptData);

                    // 自動填入數量到列印設定（忽略小數點）
                    string processedQuantity = ProcessQuantityValue(_currentReceiptData.Quantity.ToString());
                    textBoxQuantity.Text = processedQuantity;

                    // 檢查是否已載入 BTW 檔案
                    if (!string.IsNullOrEmpty(textBoxBtwPath.Text) && _barTenderService != null)
                    {
                        try
                        {
                            Console.WriteLine("=== 開始填入 API 資料到 BarTender 欄位 ===");
                            Console.WriteLine($"ProductName 長度: {_currentReceiptData.ProductName?.Length ?? 0} 字元");
                            Console.WriteLine($"ProductName 內容: {_currentReceiptData.ProductName}");

                            // 策略1：調整欄位設定順序 - 先設定短文字欄位，最後設定長文字欄位
                            // 第一批：基本短文字欄位
                            _barTenderService.SetFieldValue("DocType", _currentReceiptData.DocType);
                            _barTenderService.SetFieldValue("DocNumber", _currentReceiptData.DocNumber);
                            _barTenderService.SetFieldValue("DocItem", _currentReceiptData.DocItem);
                            _barTenderService.SetFieldValue("Date", _currentReceiptData.PurchaseDate);
                            _barTenderService.SetFieldValue("ProductNumber", _currentReceiptData.ProductCode);

                            // 第二批：中等長度文字欄位
                            _barTenderService.SetFieldValue("SupplierName", _currentReceiptData.SupplierName);

                            // 第三批：最後設定長文字欄位（ProductName）
                            _barTenderService.SetFieldValue("ProductName", _currentReceiptData.ProductName);

                            // 月份從日期中提取 - 支援多種日期格式
                            string monthValue = ExtractMonthFromDate(_currentReceiptData.PurchaseDate);
                            _barTenderService.SetFieldValue("Months", monthValue);
                            Console.WriteLine($"從日期 {_currentReceiptData.PurchaseDate} 提取月份: {monthValue}");

                            // 設定使用者輸入的欄位（保持當前值或使用預設值，忽略小數點）
                            string processedQuantity2 = ProcessQuantityValue(textBoxQuantity.Text);
                            _barTenderService.SetFieldValue("Quantity", processedQuantity2);
                            _barTenderService.SetFieldValue("DC", string.IsNullOrEmpty(textBoxDc.Text) ? "" : textBoxDc.Text);
                            _barTenderService.SetFieldValue("HWVer", string.IsNullOrEmpty(textBoxHwVer.Text) ? "" : textBoxHwVer.Text);
                            _barTenderService.SetFieldValue("FWVer", string.IsNullOrEmpty(textBoxFwVer.Text) ? "" : textBoxFwVer.Text);

                            Console.WriteLine("=== API 資料填入完成 ===");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"填入 API 資料時發生錯誤: {ex.Message}");
                            MessageBox.Show($"填入資料時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        Console.WriteLine("未載入 BTW 檔案，跳過預覽更新");
                    }

                    // 步驟3：API 資料獲取成功，開啟列印設定
                    SetStep3State();

                    // 更新 ProductName 顯示
                    UpdateProductNameDisplay();

                    // 在 ProductName 更新後才刷新預覽
                    if (!string.IsNullOrEmpty(textBoxBtwPath.Text) && _barTenderService != null)
                    {
                        Console.WriteLine("=== 開始刷新預覽（包含 ProductName）===");
                        ShowLabelPreview();
                        Console.WriteLine("=== 預覽刷新完成 ===");
                    }

                    MessageBox.Show("資料獲取成功！預覽已更新。\n列印設定功能已開啟。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    Console.WriteLine("API 資料獲取成功，已開啟列印設定功能");
                }
                else
                {
                    richTextBoxApiResult.Text = "查無相關資料，請確認輸入的單據資訊是否正確";

                    string docCategory = comboBoxDocCategory.SelectedIndex == 1 ? "托外進貨" : "一般進貨";
                    MessageBox.Show($"查無相關資料！\n\n請確認以下資訊是否正確：\n" +
                                  $"• 單據類型：{docCategory}\n" +
                                  $"• 進貨單別：{textBoxDocType.Text}\n" +
                                  $"• 進貨單號：{textBoxDocNumber.Text}\n" +
                                  $"• 進貨項次：{textBoxDocItem.Text}\n\n" +
                                  "請重新輸入正確的單據資訊。",
                                  "資料查詢失敗",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Warning);

                    // API 獲取失敗，保持在步驟2
                    Console.WriteLine("API 資料獲取失敗，保持在步驟2狀態");
                }
            }
            catch (Exception ex)
            {
                richTextBoxApiResult.Text = $"錯誤: {ex.Message}";
                MessageBox.Show($"獲取資料時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢復按鈕狀態
                buttonGetData.Enabled = true;
                buttonGetData.Text = "獲取資料";
            }
        }

        /// <summary>
        /// 顯示 API 回傳結果
        /// </summary>
        private void DisplayApiResult(PurchaseReceiptModel data)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"進貨單別: {data.DocType}");
            sb.AppendLine($"進貨單號: {data.DocNumber}");
            sb.AppendLine($"進貨項次: {data.DocItem}");
            sb.AppendLine($"進貨日期: {data.PurchaseDate}");
            sb.AppendLine($"供應廠商: {data.SupplierName} ({data.SupplierCode})");
            sb.AppendLine($"品號: {data.ProductCode}");
            sb.AppendLine($"品名: {data.ProductName}");
            sb.AppendLine($"規格: {data.Specification}");
            sb.AppendLine($"進貨數量: {data.Quantity}");

            richTextBoxApiResult.Text = sb.ToString();
        }



        /// <summary>
        /// 更新預覽按鈕點擊事件
        /// </summary>
        private void buttonUpdatePreview_Click(object sender, EventArgs e)
        {
            try
            {
                // 驗證是否已載入 BTW 檔案
                if (string.IsNullOrWhiteSpace(textBoxBtwPath.Text))
                {
                    MessageBox.Show("請先選擇 BTW 檔案", "更新預覽", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 禁用按鈕防止重複點擊
                buttonUpdatePreview.Enabled = false;
                buttonUpdatePreview.Text = "更新中...";

                Console.WriteLine("=== 手動觸發預覽更新 ===");

                // 如果有 API 資料，使用 API 資料；否則使用當前輸入值
                if (_currentReceiptData != null)
                {
                    Console.WriteLine("使用 API 資料更新預覽");
                    // 重新設定所有欄位值（包含使用者修改的列印設定）
                    SetBarTenderFields();
                }
                else
                {
                    Console.WriteLine("使用當前輸入值更新預覽");
                    // 只使用使用者輸入的列印設定值
                    UpdatePreviewWithCurrentSettings();
                }

                // 更新預覽
                ShowLabelPreview();

                Console.WriteLine("=== 手動預覽更新完成 ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"手動更新預覽失敗: {ex.Message}");
                MessageBox.Show($"更新預覽時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢復按鈕狀態
                buttonUpdatePreview.Enabled = true;
                buttonUpdatePreview.Text = "更新預覽";
            }
        }

        /// <summary>
        /// 使用當前設定更新預覽（沒有 API 資料時）
        /// </summary>
        private void UpdatePreviewWithCurrentSettings()
        {
            try
            {
                // 處理數量（忽略小數點）
                string processedQuantity = ProcessQuantityValue(textBoxQuantity.Text);
                _barTenderService.SetFieldValue("Quantity", processedQuantity);

                // 設定其他使用者輸入的欄位
                _barTenderService.SetFieldValue("DC", textBoxDc.Text);
                _barTenderService.SetFieldValue("HWVer", textBoxHwVer.Text);
                _barTenderService.SetFieldValue("FWVer", textBoxFwVer.Text);

                Console.WriteLine($"已更新列印設定: 數量={processedQuantity}, DC={textBoxDc.Text}, HW Ver={textBoxHwVer.Text}, FW Ver={textBoxFwVer.Text}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新當前設定時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 顯示標籤預覽
        /// </summary>
        private void ShowLabelPreview()
        {
            // 使用鎖定機制避免併發存取
            lock (_previewLock)
            {
                Image oldImage = null;
                try
                {
                    Console.WriteLine("=== 開始生成標籤預覽 ===");

                    // 優先使用 API 資料，如果沒有則使用示例資料
                    if (_currentReceiptData != null)
                    {
                        Console.WriteLine("使用 API 資料生成預覽");
                        SetBarTenderFields();
                    }
                    else
                    {
                        Console.WriteLine("使用示例資料生成預覽");
                        // 設定預設示例資料
                        _barTenderService.SetFieldValue("SupplierName", "XX科技(股)公司");
                        _barTenderService.SetFieldValue("Months", "12");
                        _barTenderService.SetFieldValue("ProductName", "FPC-7602 BOTTOM COVER BRACKET");
                        _barTenderService.SetFieldValue("DocType", "3113");
                        _barTenderService.SetFieldValue("DocItem", "0001");
                        _barTenderService.SetFieldValue("DocNumber", "1140421033");
                        _barTenderService.SetFieldValue("ProductNumber", "371760236010SP");
                        _barTenderService.SetFieldValue("Quantity", "1000");
                        _barTenderService.SetFieldValue("DC", "250601");
                        _barTenderService.SetFieldValue("HWVer", "V1.01");
                        _barTenderService.SetFieldValue("FWVer", "V1.6F");
                        _barTenderService.SetFieldValue("Date", "2025/06/06");
                    }

                    Console.WriteLine("開始獲取預覽圖片...");

                    // 獲取預覽圖片
                    Image previewImage = _barTenderService.GetLabelPreview();

                    // 安全地更新 PictureBox
                    if (this.InvokeRequired)
                    {
                        this.Invoke(new Action(() => UpdatePictureBoxSafely(previewImage, ref oldImage)));
                    }
                    else
                    {
                        UpdatePictureBoxSafely(previewImage, ref oldImage);
                    }

                    Console.WriteLine("=== 標籤預覽生成完成 ===");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"顯示預覽時發生錯誤：{ex.Message}");
                    Console.WriteLine($"錯誤堆疊: {ex.StackTrace}");

                    // 安全地設定錯誤狀態
                    if (this.InvokeRequired)
                    {
                        this.Invoke(new Action(() => SetPreviewErrorState()));
                    }
                    else
                    {
                        SetPreviewErrorState();
                    }

                    MessageBox.Show($"顯示預覽時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // 確保舊圖片資源被釋放
                    if (oldImage != null)
                    {
                        try
                        {
                            oldImage.Dispose();
                            Console.WriteLine("已釋放舊預覽圖片資源");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"釋放舊圖片資源時發生錯誤: {ex.Message}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 安全地更新 PictureBox 圖片
        /// </summary>
        private void UpdatePictureBoxSafely(Image newImage, ref Image oldImage)
        {
            try
            {
                // 保存舊圖片的參考以便稍後釋放
                oldImage = pictureBoxPreview.Image;

                if (newImage != null)
                {
                    pictureBoxPreview.Image = newImage;
                    pictureBoxPreview.BackColor = Color.White;
                    Console.WriteLine($"預覽圖片已更新，尺寸: {newImage.Width}x{newImage.Height}");
                }
                else
                {
                    pictureBoxPreview.Image = null;
                    pictureBoxPreview.BackColor = Color.LightGray;
                    Console.WriteLine("預覽圖片生成失敗");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新 PictureBox 時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 設定預覽錯誤狀態
        /// </summary>
        private void SetPreviewErrorState()
        {
            try
            {
                if (pictureBoxPreview.Image != null)
                {
                    pictureBoxPreview.Image.Dispose();
                }
                pictureBoxPreview.Image = null;
                pictureBoxPreview.BackColor = Color.LightCoral;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"設定預覽錯誤狀態時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// 設定 BarTender 欄位值
        /// </summary>
        private void SetBarTenderFields()
        {
            if (_currentReceiptData == null) return;

            try
            {
                // 策略1：調整欄位設定順序 - 先設定短文字欄位，最後設定長文字欄位
                // 優先使用 textBoxProductName.Text，如果為空則使用 API 資料
                string currentProductName = !string.IsNullOrEmpty(textBoxProductName.Text)
                    ? textBoxProductName.Text
                    : _currentReceiptData.ProductName;
                Console.WriteLine($"手動更新預覽 - ProductName 長度: {currentProductName?.Length ?? 0} 字元");

                // 第一批：基本短文字欄位
                _barTenderService.SetFieldValue("DocType", _currentReceiptData.DocType);
                _barTenderService.SetFieldValue("DocNumber", _currentReceiptData.DocNumber);
                _barTenderService.SetFieldValue("DocItem", _currentReceiptData.DocItem);
                _barTenderService.SetFieldValue("Date", _currentReceiptData.PurchaseDate);
                _barTenderService.SetFieldValue("ProductNumber", _currentReceiptData.ProductCode);

                // 月份從日期中提取
                string monthValue = ExtractMonthFromDate(_currentReceiptData.PurchaseDate);
                _barTenderService.SetFieldValue("Months", monthValue);

                // 第二批：中等長度文字欄位
                _barTenderService.SetFieldValue("SupplierName", _currentReceiptData.SupplierName);

                // 設定使用者輸入的欄位（忽略小數點）
                string processedQuantity = ProcessQuantityValue(textBoxQuantity.Text);
                _barTenderService.SetFieldValue("Quantity", processedQuantity);
                _barTenderService.SetFieldValue("DC", textBoxDc.Text);
                _barTenderService.SetFieldValue("HWVer", textBoxHwVer.Text);
                _barTenderService.SetFieldValue("FWVer", textBoxFwVer.Text);

                // 第三批：最後設定長文字欄位（ProductName）- 使用使用者可能修改過的版本
                _barTenderService.SetFieldValue("ProductName", currentProductName);
                Console.WriteLine($"使用 ProductName: {currentProductName} (長度: {currentProductName?.Length ?? 0})");
            }
            catch (Exception ex)
            {
                throw new Exception($"設定標籤欄位時發生錯誤: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 更新 ProductName 顯示
        /// </summary>
        private void UpdateProductNameDisplay()
        {
            if (_currentReceiptData != null)
            {
                string productName = _currentReceiptData.ProductName ?? "";
                textBoxProductName.Text = productName;
                UpdateProductNameLength();

                Console.WriteLine($"更新 ProductName 顯示: {productName}");
                Console.WriteLine($"ProductName 長度: {productName.Length} 字元");
            }
        }

        /// <summary>
        /// 更新 ProductName 長度顯示
        /// </summary>
        private void UpdateProductNameLength()
        {
            int length = textBoxProductName.Text?.Length ?? 0;
            labelProductNameLength.Text = $"長度: {length}";

            // 根據長度設定顏色提示
            if (length > 50)
            {
                labelProductNameLength.ForeColor = Color.Red;
                labelProductNameLength.Text += " (過長，建議縮短)";
            }
            else if (length > 30)
            {
                labelProductNameLength.ForeColor = Color.Orange;
                labelProductNameLength.Text += " (較長)";
            }
            else
            {
                labelProductNameLength.ForeColor = Color.Gray;
            }
        }

        /// <summary>
        /// ProductName TextBox 文字變更事件
        /// </summary>
        private void textBoxProductName_TextChanged(object sender, EventArgs e)
        {
            UpdateProductNameLength();

            // 如果有 API 資料，更新 ProductName
            if (_currentReceiptData != null)
            {
                _currentReceiptData.ProductName = textBoxProductName.Text;
                Console.WriteLine($"使用者修改 ProductName: {textBoxProductName.Text} (長度: {textBoxProductName.Text.Length})");
            }
        }

        /// <summary>
        /// 從日期字串中提取月份
        /// </summary>
        /// <param name="dateString">日期字串</param>
        /// <returns>月份字串</returns>
        private string ExtractMonthFromDate(string dateString)
        {
            if (string.IsNullOrEmpty(dateString))
            {
                Console.WriteLine("日期字串為空，使用預設月份: 12");
                return "12";
            }

            // 支援的日期格式
            string[] dateFormats = {
                "yyyyMMdd",        // 20250602 格式
                "yyyy/MM/dd",
                "yyyy-MM-dd",
                "yyyy/M/d",
                "yyyy-M-d",
                "MM/dd/yyyy",
                "M/d/yyyy",
                "dd/MM/yyyy",
                "d/M/yyyy",
                "yyyy年MM月dd日",
                "yyyy年M月d日"
            };

            foreach (string format in dateFormats)
            {
                if (DateTime.TryParseExact(dateString, format, null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                {
                    Console.WriteLine($"使用格式 '{format}' 成功解析日期: {dateString} -> 月份: {parsedDate.Month}");
                    return parsedDate.Month.ToString();
                }
            }

            // 如果精確格式解析失敗，嘗試通用解析
            if (DateTime.TryParse(dateString, out DateTime generalParsedDate))
            {
                Console.WriteLine($"使用通用格式成功解析日期: {dateString} -> 月份: {generalParsedDate.Month}");
                return generalParsedDate.Month.ToString();
            }

            // 嘗試從字串中提取數字（如果是特殊格式）
            var match = System.Text.RegularExpressions.Regex.Match(dateString, @"(\d{1,2})[/\-年月]");
            if (match.Success)
            {
                if (int.TryParse(match.Groups[1].Value, out int month) && month >= 1 && month <= 12)
                {
                    Console.WriteLine($"使用正則表達式從日期 '{dateString}' 提取月份: {month}");
                    return month.ToString();
                }
            }

            Console.WriteLine($"無法解析日期 '{dateString}'，使用預設月份: 12");
            return "12"; // 預設值
        }

        /// <summary>
        /// 處理數量值，忽略小數點後的數字
        /// </summary>
        /// <param name="quantityString">數量字串</param>
        /// <returns>整數數量字串</returns>
        private string ProcessQuantityValue(string quantityString)
        {
            if (string.IsNullOrEmpty(quantityString))
            {
                Console.WriteLine("數量字串為空，使用預設值: 1");
                return "1";
            }

            // 嘗試解析為 decimal 以處理小數
            if (decimal.TryParse(quantityString, out decimal quantity))
            {
                // 取整數部分
                int integerQuantity = (int)Math.Floor(quantity);
                Console.WriteLine($"數量處理: '{quantityString}' -> {integerQuantity} (忽略小數部分)");
                return integerQuantity.ToString();
            }

            // 如果解析失敗，嘗試提取數字部分
            var match = System.Text.RegularExpressions.Regex.Match(quantityString, @"(\d+)");
            if (match.Success)
            {
                Console.WriteLine($"從字串 '{quantityString}' 提取數量: {match.Groups[1].Value}");
                return match.Groups[1].Value;
            }

            Console.WriteLine($"無法解析數量 '{quantityString}'，使用預設值: 1");
            return "1"; // 預設值
        }

        /// <summary>
        /// 初始化印表機清單
        /// </summary>
        private void InitializePrinters()
        {
            try
            {
                if (_barTenderService != null)
                {
                    // 獲取可用印表機清單
                    List<string> printers = _barTenderService.GetAvailablePrinters();

                    comboBoxPrinter.Items.Clear();
                    foreach (string printer in printers)
                    {
                        comboBoxPrinter.Items.Add(printer);
                    }

                    // 設定預設印表機
                    if (comboBoxPrinter.Items.Count > 0)
                    {
                        string defaultPrinter = _barTenderService.GetDefaultPrinter();
                        int defaultIndex = comboBoxPrinter.Items.IndexOf(defaultPrinter);

                        if (defaultIndex >= 0)
                        {
                            comboBoxPrinter.SelectedIndex = defaultIndex;
                        }
                        else
                        {
                            comboBoxPrinter.SelectedIndex = 0;
                        }

                        // 設定選擇的印表機到服務
                        _barTenderService.SetSelectedPrinter(comboBoxPrinter.SelectedItem.ToString());
                    }

                    Console.WriteLine($"印表機初始化完成，共找到 {printers.Count} 台印表機");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"初始化印表機時發生錯誤: {ex.Message}");
                MessageBox.Show($"初始化印表機時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 印表機選擇變更事件
        /// </summary>
        private void comboBoxPrinter_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBoxPrinter.SelectedItem != null && _barTenderService != null)
                {
                    string selectedPrinter = comboBoxPrinter.SelectedItem.ToString();
                    _barTenderService.SetSelectedPrinter(selectedPrinter);
                    Console.WriteLine($"選擇印表機: {selectedPrinter}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"設定印表機時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// 列印按鈕點擊事件
        /// </summary>
        private void buttonPrint_Click(object sender, EventArgs e)
        {
            try
            {
                // 驗證列印份數
                if (!int.TryParse(textBoxCopies.Text, out int copies) || copies < 1 || copies > 999)
                {
                    MessageBox.Show("請輸入有效的列印份數（1-999）", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBoxCopies.Focus();
                    return;
                }

                // 確認印表機選擇
                if (comboBoxPrinter.SelectedItem == null)
                {
                    MessageBox.Show("請選擇印表機", "選擇印表機", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 確認是否要列印
                string printerName = comboBoxPrinter.SelectedItem.ToString();
                DialogResult result = MessageBox.Show(
                    $"確定要列印 {copies} 份標籤到印表機「{printerName}」嗎？",
                    "確認列印",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Console.WriteLine($"開始列印標籤: 印表機={printerName}, 份數={copies}");

                    // 🎯 關鍵修復：列印前確保所有欄位值都是最新的
                    Console.WriteLine("=== 列印前最終欄位同步 ===");
                    try
                    {
                        if (_currentReceiptData != null)
                        {
                            Console.WriteLine("使用 API 資料進行最終同步");
                            // 重新設定所有欄位值（包含使用者修改的列印設定）
                            SetBarTenderFields();
                        }
                        else
                        {
                            Console.WriteLine("使用當前輸入值進行最終同步");
                            // 只使用使用者輸入的列印設定值
                            UpdatePreviewWithCurrentSettings();
                        }
                        Console.WriteLine("列印前欄位同步完成");
                    }
                    catch (Exception syncEx)
                    {
                        Console.WriteLine($"列印前欄位同步發生錯誤: {syncEx.Message}");
                        MessageBox.Show($"列印前欄位同步發生錯誤: {syncEx.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return; // 如果同步失敗，不執行列印
                    }

                    // 使用 BackgroundWorker 執行列印
                    StartPrintJob(printerName, copies);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"列印標籤時發生錯誤: {ex.Message}");
                MessageBox.Show($"列印標籤時發生錯誤: {ex.Message}", "列印錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 開始列印工作
        /// </summary>
        private void StartPrintJob(string printerName, int copies)
        {
            try
            {
                // 禁用列印按鈕防止重複點擊
                buttonPrint.Enabled = false;

                // 顯示進度對話框
                _progressForm = new ProgressForm("正在發送列印工作...");

                // 設定 BackgroundWorker
                _printWorker = new BackgroundWorker();
                _printWorker.WorkerReportsProgress = true;
                _printWorker.WorkerSupportsCancellation = true;
                _printWorker.DoWork += PrintWorker_DoWork;
                _printWorker.ProgressChanged += PrintWorker_ProgressChanged;
                _printWorker.RunWorkerCompleted += PrintWorker_RunWorkerCompleted;

                // 準備列印資料
                var printData = new BarTenderService.PrintJobData
                {
                    PrinterName = printerName,
                    Copies = copies
                };

                // 顯示進度對話框（模態）
                _progressForm.Show(this);

                // 開始背景工作
                _printWorker.RunWorkerAsync(printData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"啟動列印工作時發生錯誤: {ex.Message}");
                MessageBox.Show($"啟動列印工作時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 恢復按鈕狀態
                buttonPrint.Enabled = true;

                // 關閉進度對話框
                _progressForm?.Close();
            }
        }

        /// <summary>
        /// BackgroundWorker 執行工作事件
        /// </summary>
        private void PrintWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var printData = e.Argument as BarTenderService.PrintJobData;

            try
            {
                // 設定超時時間（30秒）
                var startTime = DateTime.Now;
                var timeout = TimeSpan.FromSeconds(30);

                // 執行列印工作
                var result = _barTenderService.PrintLabelBackground(printData, worker);

                // 檢查是否超時
                if (DateTime.Now - startTime > timeout)
                {
                    result.Success = false;
                    result.Message = "列印工作超時";
                }

                e.Result = result;
            }
            catch (Exception ex)
            {
                var errorResult = new BarTenderService.PrintJobResult
                {
                    Success = false,
                    Message = $"列印工作發生異常: {ex.Message}",
                    Error = ex
                };
                e.Result = errorResult;
            }
        }

        /// <summary>
        /// BackgroundWorker 進度更新事件
        /// </summary>
        private void PrintWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    string message = e.UserState?.ToString() ?? "處理中...";
                    _progressForm.SetMessage(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新進度時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// BackgroundWorker 完成事件
        /// </summary>
        private void PrintWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                // 關閉進度對話框
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    _progressForm.Close();
                    _progressForm = null;
                }

                // 恢復按鈕狀態
                buttonPrint.Enabled = true;

                // 處理結果
                if (e.Error != null)
                {
                    Console.WriteLine($"列印工作發生錯誤: {e.Error.Message}");
                    MessageBox.Show($"列印工作發生錯誤: {e.Error.Message}", "列印錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (e.Cancelled)
                {
                    Console.WriteLine("列印工作已取消");
                    MessageBox.Show("列印工作已取消", "已取消", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    var result = e.Result as BarTenderService.PrintJobResult;
                    if (result != null)
                    {
                        if (result.Success)
                        {
                            Console.WriteLine("列印工作完成");
                            MessageBox.Show($"✅ 已傳送至印表機\n\n" +
                                          $"印表機：{_barTenderService.GetSelectedPrinter()}\n" +
                                          $"位置：{result.PrinterLocation}\n" +
                                          $"份數：{result.Copies}\n\n" +
                                          "列印工作已發送，請稍候取件。",
                                          "列印已發送",
                                          MessageBoxButtons.OK,
                                          MessageBoxIcon.Information);
                        }
                        else
                        {
                            Console.WriteLine($"列印工作失敗: {result.Message}");
                            MessageBox.Show($"列印失敗: {result.Message}", "列印失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                // 清理 BackgroundWorker
                if (_printWorker != null)
                {
                    _printWorker.Dispose();
                    _printWorker = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"處理列印完成事件時發生錯誤: {ex.Message}");
                MessageBox.Show($"處理列印結果時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 表單關閉時釋放資源
        /// </summary>
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                // 取消並清理 BackgroundWorker
                if (_printWorker != null && _printWorker.IsBusy)
                {
                    _printWorker.CancelAsync();
                }
                _printWorker?.Dispose();

                // 關閉進度對話框
                _progressForm?.Close();

                // 釋放其他資源
                _apiService?.Dispose();
                _barTenderService?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"釋放資源時發生錯誤: {ex.Message}");
            }
            finally
            {
                base.OnFormClosed(e);
            }
        }
    }
}
