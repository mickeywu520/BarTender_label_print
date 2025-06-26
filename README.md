# 標籤列印管理程式 (Label Print Manager)

這是一個整合 BarTender SDK 和 REST API 的標籤列印管理程式，用於自動填入 BTW 檔案的欄位並進行列印。

## 功能特色

- **BTW 檔案載入**: 支援選擇和載入 BarTender 標籤格式檔案
- **REST API 整合**: 自動從 ERP 系統獲取進貨單資料
- **智能路由**: 根據單據類型自動選擇對應的 API 端點
- **欄位自動填入**: 將 API 資料自動填入標籤欄位
- **使用者自訂欄位**: 支援手動輸入 D/C、HW Ver、FW Ver 等欄位
- **列印確認**: 提供列印前的確認對話框

## 系統需求

- .NET Framework 4.7.2 或更高版本
- BarTender SDK (需要另外安裝)
- 網路連線 (用於 API 呼叫)

## 安裝說明

1. 確保已安裝 BarTender 軟體和 SDK
2. 編譯專案前需要添加 BarTender SDK 的 COM 參考
3. 確認印表機 IP 位址設定正確 (預設: 192.168.6.47)

## 使用方式

### 1. 選擇 BTW 檔案
- 點擊「瀏覽...」按鈕選擇要使用的 BTW 檔案
- 選擇成功後會在右側顯示載入狀態

### 2. 輸入進貨單資訊
- **進貨單別**: 輸入單據類型
- **單據類型**: 從下拉選單選擇「一般進貨」或「托外進貨」
- **進貨單號**: 輸入進貨單號碼
- **進貨項次**: 輸入項次序號

### 3. 獲取資料
- 點擊「獲取資料」按鈕從 API 獲取進貨單資料
- 系統會根據下拉選單的選擇呼叫對應的 API 端點:
  - 一般進貨: `getPurchaseReceipt`
  - 托外進貨: `getSubcontractReceipt`

### 4. 設定列印參數
- **數量**: 自動填入 API 獲取的數量，可手動修改
- **D/C**: 手動輸入
- **HW Ver**: 手動輸入硬體版本
- **FW Ver**: 手動輸入韌體版本

### 5. 列印標籤
- 點擊「列印標籤」按鈕
- 確認列印資訊後執行列印

## API 設定

### 基礎 URL
```
http://192.168.0.13:100/api/ErpToBarcode/
```

### API 端點
- **一般進貨**: `getPurchaseReceipt`
- **托外進貨**: `getSubcontractReceipt`

### 請求範例
```
http://192.168.0.13:100/api/ErpToBarcode/getPurchaseReceipt?DocType=3421&DocNumber=114060001&DocItem=0001
```

### 回應格式
```json
[
  {
    "進貨單別": "3421",
    "進貨單號": "114060001",
    "進貨項次(序號)": "0001",
    "進貨日期": "20250602",
    "供應廠商代號": "LKP001",
    "供應廠商名稱": "崑璞",
    "品號": "3412182000540P",
    "品名": "CBL,POWER,QUEEN PUO,10100032004033(20100032004033)",
    "規格": "1830mm,IEC320 C13 TO EU TYPE,CEE7-WIII,W/ENEC",
    "進貨數量": 300.000
  }
]
```

## BarTender 欄位對應

程式會自動設定以下 BarTender 欄位:

| 欄位名稱 | 資料來源 | 說明 |
|---------|---------|------|
| DocType | API | 進貨單別 |
| DocNumber | API | 進貨單號 |
| DocItem | API | 進貨項次 |
| PurchaseDate | API | 進貨日期 |
| SupplierCode | API | 供應廠商代號 |
| SupplierName | API | 供應廠商名稱 |
| ProductCode | API | 品號 |
| ProductName | API | 品名 |
| Specification | API | 規格 |
| Quantity | 使用者輸入 | 列印數量 |
| DC | 使用者輸入 | D/C 欄位 |
| HWVer | 使用者輸入 | 硬體版本 |
| FWVer | 使用者輸入 | 韌體版本 |

## 注意事項

1. **BarTender SDK**: 目前程式中的 BarTender 相關程式碼已註解，需要安裝 BarTender SDK 後取消註解並添加適當的參考
2. **印表機設定**: 印表機 IP 位址預設為 192.168.6.47，如需修改請編輯 `BarTenderService.cs` 中的 `PRINTER_NAME` 常數
3. **API 連線**: 確保能夠連線到 192.168.0.13:100 的 API 服務
4. **錯誤處理**: 程式包含完整的錯誤處理機制，會顯示詳細的錯誤訊息

## 開發說明

### 專案結構
```
LabelPrintManager/
├── Models/
│   └── PurchaseReceiptModel.cs    # API 資料模型
├── Services/
│   ├── ApiService.cs              # API 服務類別
│   └── BarTenderService.cs        # BarTender SDK 服務類別
├── Form1.cs                       # 主要表單邏輯
├── Form1.Designer.cs              # 表單設計檔案
└── Form1.resx                     # 表單資源檔案
```

### 擴展功能
- 可以修改 `ApiService.cs` 來支援更多的 API 端點
- 可以在 `BarTenderService.cs` 中添加更多的標籤操作功能
- 可以擴展 `PurchaseReceiptModel.cs` 來支援更多的資料欄位

## 授權

此專案僅供內部使用。
