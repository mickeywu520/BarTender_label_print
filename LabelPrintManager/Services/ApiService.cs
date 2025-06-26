using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using LabelPrintManager.Models;

namespace LabelPrintManager.Services
{
    /// <summary>
    /// API 服務類別，處理與 ERP 系統的 REST API 通訊
    /// </summary>
    public class ApiService
    {
        private readonly HttpClient _httpClient;
        private const string BASE_URL = "http://192.168.0.13:100/api/ErpToBarcode/";

        public ApiService()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        /// <summary>
        /// 根據單據類型獲取進貨單據資料
        /// </summary>
        /// <param name="docType">進貨單別</param>
        /// <param name="docNumber">進貨單號</param>
        /// <param name="docItem">進貨項次</param>
        /// <param name="isSubcontract">是否為托外進貨</param>
        /// <returns>進貨單據資料</returns>
        public async Task<PurchaseReceiptModel> GetPurchaseReceiptAsync(string docType, string docNumber, string docItem, bool isSubcontract = false)
        {
            try
            {
                string endpoint = isSubcontract ? "getSubcontractReceipt" : "getPurchaseReceipt";
                string url = $"{BASE_URL}{endpoint}?DocType={docType}&DocNumber={docNumber}&DocItem={docItem}";

                Console.WriteLine("=== API 呼叫詳細資訊 ===");
                Console.WriteLine($"呼叫類型: {(isSubcontract ? "托外進貨" : "一般進貨")}");
                Console.WriteLine($"API 端點: {endpoint}");
                Console.WriteLine($"完整 URL: {url}");
                Console.WriteLine($"參數 - DocType: {docType}");
                Console.WriteLine($"參數 - DocNumber: {docNumber}");
                Console.WriteLine($"參數 - DocItem: {docItem}");
                Console.WriteLine($"參數 - isSubcontract: {isSubcontract}");
                Console.WriteLine("=== 開始 HTTP 請求 ===");

                HttpResponseMessage response = await _httpClient.GetAsync(url);

                Console.WriteLine($"HTTP 狀態碼: {response.StatusCode}");
                Console.WriteLine($"HTTP 狀態描述: {response.ReasonPhrase}");

                response.EnsureSuccessStatusCode();

                string jsonContent = await response.Content.ReadAsStringAsync();

                Console.WriteLine($"回應內容長度: {jsonContent?.Length ?? 0} 字元");
                Console.WriteLine($"回應內容預覽: {(jsonContent?.Length > 200 ? jsonContent.Substring(0, 200) + "..." : jsonContent)}");

                // 解析 JSON 回應
                var receipt = ParseJsonToModel(jsonContent, isSubcontract);

                if (receipt != null)
                {
                    Console.WriteLine("=== API 呼叫成功完成 ===");
                    return receipt;
                }
                else
                {
                    Console.WriteLine("=== API 回應無資料或解析失敗 ===");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("=== API 呼叫失敗 ===");
                Console.WriteLine($"錯誤類型: {ex.GetType().Name}");
                Console.WriteLine($"錯誤訊息: {ex.Message}");
                Console.WriteLine($"錯誤堆疊: {ex.StackTrace}");
                throw new Exception($"API 呼叫失敗: {ex.Message}", ex);
            }
        }



        /// <summary>
        /// 簡單的 JSON 解析方法 (暫時實作)
        /// </summary>
        /// <param name="jsonContent">JSON 字串</param>
        /// <param name="isSubcontract">是否為托外進貨</param>
        /// <returns>解析後的資料模型</returns>
        private PurchaseReceiptModel ParseJsonToModel(string jsonContent, bool isSubcontract)
        {
            try
            {
                Console.WriteLine("=== 開始解析 JSON 資料 ===");
                Console.WriteLine($"解析模式: {(isSubcontract ? "托外進貨" : "一般進貨")}");
                Console.WriteLine($"JSON 內容: {jsonContent}");

                // 檢查 JSON 回應格式
                if (string.IsNullOrEmpty(jsonContent))
                {
                    Console.WriteLine("JSON 內容為空");
                    return null;
                }

                // 簡單檢查是否為標準回應格式 {"Code":"200","Message":"Success","Data":"..."}
                if (jsonContent.Contains("\"Code\"") && jsonContent.Contains("\"Data\""))
                {
                    Console.WriteLine("檢測到標準 API 回應格式");

                    // 檢查 Data 欄位是否為空
                    if (jsonContent.Contains("\"Data\":\"[]\"") ||
                        jsonContent.Contains("\"Data\":[]") ||
                        jsonContent.Contains("\"Data\":\"\"") ||
                        jsonContent.Contains("\"Data\":null"))
                    {
                        Console.WriteLine("API 回應 Data 欄位為空，表示查無資料");
                        return null; // 返回 null 表示查無資料
                    }

                    // 檢查是否為成功回應
                    if (!jsonContent.Contains("\"Code\":\"200\""))
                    {
                        Console.WriteLine("API 回應非成功狀態");
                        return null;
                    }

                    // 有實際資料，進行真正的 JSON 解析
                    Console.WriteLine("API 回應包含資料，開始解析真實 JSON");

                    // 解析真實的 JSON 資料
                    return ParseRealJsonData(jsonContent, isSubcontract);
                }
                else
                {
                    // 非標準格式，可能是其他格式或錯誤
                    Console.WriteLine("非標準 API 回應格式，無法解析");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解析 JSON 時發生錯誤: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 解析真實的 JSON 資料
        /// </summary>
        /// <param name="jsonContent">JSON 字串</param>
        /// <param name="isSubcontract">是否為托外進貨</param>
        /// <returns>解析後的資料模型</returns>
        private PurchaseReceiptModel ParseRealJsonData(string jsonContent, bool isSubcontract)
        {
            try
            {
                Console.WriteLine("=== 開始解析真實 JSON 資料 ===");

                // 簡單的字串解析方法（避免引入額外的 JSON 庫）
                // 從 Data 欄位中提取實際的 JSON 陣列

                // 找到 "Data":" 的位置
                int dataStart = jsonContent.IndexOf("\"Data\":\"");
                if (dataStart == -1)
                {
                    Console.WriteLine("找不到 Data 欄位");
                    return null;
                }

                // 移動到實際資料開始位置
                dataStart += 8; // "Data":"的長度

                // 找到資料結束位置（最後一個 "）
                int dataEnd = jsonContent.LastIndexOf("\"");
                if (dataEnd <= dataStart)
                {
                    Console.WriteLine("找不到 Data 欄位結束位置");
                    return null;
                }

                // 提取 Data 內容並處理轉義字符
                string dataContent = jsonContent.Substring(dataStart, dataEnd - dataStart);
                dataContent = dataContent.Replace("\\r\\n", "").Replace("\\\"", "\"");

                Console.WriteLine($"提取的 Data 內容: {dataContent}");

                // 解析資料欄位
                var model = new PurchaseReceiptModel();

                if (isSubcontract)
                {
                    // 托外進貨欄位對應
                    model.DocType = ExtractJsonValue(dataContent, "託外進貨單別");
                    model.DocNumber = ExtractJsonValue(dataContent, "託外進貨單號").Trim();
                    model.DocItem = ExtractJsonValue(dataContent, "託外進貨項次(序號)");
                    model.PurchaseDate = ExtractJsonValue(dataContent, "託外進貨日期");
                    model.SupplierCode = ExtractJsonValue(dataContent, "供應廠商代號");
                    model.SupplierName = ExtractJsonValue(dataContent, "供應廠商名稱");
                    model.ProductCode = ExtractJsonValue(dataContent, "品號");
                    model.ProductName = ExtractJsonValue(dataContent, "品名");
                    model.Specification = ExtractJsonValue(dataContent, "規格");

                    // 解析數量
                    string quantityStr = ExtractJsonValue(dataContent, "進貨數量");
                    if (decimal.TryParse(quantityStr, out decimal quantity))
                    {
                        model.Quantity = quantity;
                    }

                    Console.WriteLine("=== 托外進貨資料解析完成 ===");
                }
                else
                {
                    // 一般進貨欄位對應（根據實際 API 回應調整）
                    model.DocType = ExtractJsonValue(dataContent, "進貨單別");
                    model.DocNumber = ExtractJsonValue(dataContent, "進貨單號").Trim();
                    model.DocItem = ExtractJsonValue(dataContent, "進貨項次(序號)"); // 修正：實際 key 包含 "(序號)"
                    model.PurchaseDate = ExtractJsonValue(dataContent, "進貨日期");
                    model.SupplierCode = ExtractJsonValue(dataContent, "供應廠商代號");
                    model.SupplierName = ExtractJsonValue(dataContent, "供應廠商名稱");
                    model.ProductCode = ExtractJsonValue(dataContent, "品號");
                    model.ProductName = ExtractJsonValue(dataContent, "品名");
                    model.Specification = ExtractJsonValue(dataContent, "規格");

                    // 解析數量
                    string quantityStr = ExtractJsonValue(dataContent, "進貨數量");
                    if (decimal.TryParse(quantityStr, out decimal quantity))
                    {
                        model.Quantity = quantity;
                    }

                    Console.WriteLine("=== 一般進貨資料解析完成 ===");
                }

                // 顯示解析結果
                Console.WriteLine($"DocType: {model.DocType}");
                Console.WriteLine($"DocNumber: {model.DocNumber}");
                Console.WriteLine($"DocItem: {model.DocItem}");
                Console.WriteLine($"PurchaseDate: {model.PurchaseDate}");
                Console.WriteLine($"SupplierCode: {model.SupplierCode}");
                Console.WriteLine($"SupplierName: {model.SupplierName}");
                Console.WriteLine($"ProductCode: {model.ProductCode}");
                Console.WriteLine($"ProductName: {model.ProductName}");
                Console.WriteLine($"Specification: {model.Specification}");
                Console.WriteLine($"Quantity: {model.Quantity}");

                return model;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解析真實 JSON 資料時發生錯誤: {ex.Message}");
                Console.WriteLine($"錯誤詳細資訊: {ex}");
                return null;
            }
        }

        /// <summary>
        /// 從 JSON 字串中提取指定欄位的值
        /// </summary>
        /// <param name="jsonContent">JSON 內容</param>
        /// <param name="fieldName">欄位名稱</param>
        /// <returns>欄位值</returns>
        private string ExtractJsonValue(string jsonContent, string fieldName)
        {
            try
            {
                // 先嘗試字串值格式 "fieldName": "value"
                string searchPattern = $"\"{fieldName}\": \"";
                int startIndex = jsonContent.IndexOf(searchPattern);

                if (startIndex == -1)
                {
                    // 嘗試沒有空格的字串格式 "fieldName":"value"
                    searchPattern = $"\"{fieldName}\":\"";
                    startIndex = jsonContent.IndexOf(searchPattern);
                }

                if (startIndex != -1)
                {
                    // 找到字串值格式
                    startIndex += searchPattern.Length;
                    int endIndex = jsonContent.IndexOf("\"", startIndex);
                    if (endIndex != -1)
                    {
                        string value = jsonContent.Substring(startIndex, endIndex - startIndex);
                        Console.WriteLine($"提取欄位 {fieldName} (字串): {value}");
                        return value;
                    }
                }

                // 如果沒找到字串格式，嘗試數值格式 "fieldName": value
                searchPattern = $"\"{fieldName}\": ";
                startIndex = jsonContent.IndexOf(searchPattern);

                if (startIndex == -1)
                {
                    // 嘗試沒有空格的數值格式 "fieldName":value
                    searchPattern = $"\"{fieldName}\":";
                    startIndex = jsonContent.IndexOf(searchPattern);
                }

                if (startIndex != -1)
                {
                    // 找到數值格式
                    startIndex += searchPattern.Length;

                    // 跳過可能的空格
                    while (startIndex < jsonContent.Length && jsonContent[startIndex] == ' ')
                    {
                        startIndex++;
                    }

                    // 找到數值的結束位置（逗號、換行、或右大括號）
                    int endIndex = startIndex;
                    while (endIndex < jsonContent.Length)
                    {
                        char c = jsonContent[endIndex];
                        if (c == ',' || c == '\r' || c == '\n' || c == '}' || c == ' ')
                        {
                            break;
                        }
                        endIndex++;
                    }

                    if (endIndex > startIndex)
                    {
                        string value = jsonContent.Substring(startIndex, endIndex - startIndex).Trim();
                        Console.WriteLine($"提取欄位 {fieldName} (數值): {value}");
                        return value;
                    }
                }

                Console.WriteLine($"找不到欄位: {fieldName}");
                return "";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取欄位 {fieldName} 時發生錯誤: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 釋放資源
        /// </summary>
        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}
