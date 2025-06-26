namespace LabelPrintManager
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBoxBtwFile = new System.Windows.Forms.GroupBox();
            this.pictureBoxPreview = new System.Windows.Forms.PictureBox();
            this.buttonBrowseBtw = new System.Windows.Forms.Button();
            this.textBoxBtwPath = new System.Windows.Forms.TextBox();
            this.labelBtwFile = new System.Windows.Forms.Label();
            this.groupBoxInput = new System.Windows.Forms.GroupBox();
            this.comboBoxDocCategory = new System.Windows.Forms.ComboBox();
            this.labelDocCategory = new System.Windows.Forms.Label();
            this.buttonGetData = new System.Windows.Forms.Button();
            this.textBoxDocItem = new System.Windows.Forms.TextBox();
            this.textBoxDocNumber = new System.Windows.Forms.TextBox();
            this.textBoxDocType = new System.Windows.Forms.TextBox();
            this.labelDocItem = new System.Windows.Forms.Label();
            this.labelDocNumber = new System.Windows.Forms.Label();
            this.labelDocType = new System.Windows.Forms.Label();
            this.groupBoxApiData = new System.Windows.Forms.GroupBox();
            this.richTextBoxApiResult = new System.Windows.Forms.RichTextBox();
            this.groupBoxPrintSettings = new System.Windows.Forms.GroupBox();
            this.buttonPrint = new System.Windows.Forms.Button();
            this.buttonUpdatePreview = new System.Windows.Forms.Button();
            this.textBoxFwVer = new System.Windows.Forms.TextBox();
            this.textBoxHwVer = new System.Windows.Forms.TextBox();
            this.textBoxDc = new System.Windows.Forms.TextBox();
            this.textBoxQuantity = new System.Windows.Forms.TextBox();
            this.labelFwVer = new System.Windows.Forms.Label();
            this.labelHwVer = new System.Windows.Forms.Label();
            this.labelDc = new System.Windows.Forms.Label();
            this.labelQuantity = new System.Windows.Forms.Label();
            this.labelProductNameTitle = new System.Windows.Forms.Label();
            this.textBoxProductName = new System.Windows.Forms.TextBox();
            this.labelProductNameLength = new System.Windows.Forms.Label();
            this.openFileDialogBtw = new System.Windows.Forms.OpenFileDialog();
            this.groupBoxBtwFile.SuspendLayout();
            this.groupBoxInput.SuspendLayout();
            this.groupBoxApiData.SuspendLayout();
            this.groupBoxPrintSettings.SuspendLayout();
            this.SuspendLayout();
            //
            // groupBoxBtwFile
            //
            this.groupBoxBtwFile.Controls.Add(this.pictureBoxPreview);
            this.groupBoxBtwFile.Controls.Add(this.buttonBrowseBtw);
            this.groupBoxBtwFile.Controls.Add(this.textBoxBtwPath);
            this.groupBoxBtwFile.Controls.Add(this.labelBtwFile);
            this.groupBoxBtwFile.Location = new System.Drawing.Point(12, 12);
            this.groupBoxBtwFile.Name = "groupBoxBtwFile";
            this.groupBoxBtwFile.Size = new System.Drawing.Size(1176, 350);
            this.groupBoxBtwFile.TabIndex = 0;
            this.groupBoxBtwFile.TabStop = false;
            this.groupBoxBtwFile.Text = "BTW 檔案選擇";
            //
            // pictureBoxPreview
            //
            this.pictureBoxPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxPreview.Location = new System.Drawing.Point(400, 25);
            this.pictureBoxPreview.Name = "pictureBoxPreview";
            this.pictureBoxPreview.Size = new System.Drawing.Size(760, 310);
            this.pictureBoxPreview.TabIndex = 3;
            this.pictureBoxPreview.TabStop = false;
            this.pictureBoxPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxPreview.BackColor = System.Drawing.Color.White;
            //
            // buttonBrowseBtw
            //
            this.buttonBrowseBtw.Location = new System.Drawing.Point(310, 25);
            this.buttonBrowseBtw.Name = "buttonBrowseBtw";
            this.buttonBrowseBtw.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowseBtw.TabIndex = 2;
            this.buttonBrowseBtw.Text = "瀏覽...";
            this.buttonBrowseBtw.UseVisualStyleBackColor = true;
            this.buttonBrowseBtw.Click += new System.EventHandler(this.buttonBrowseBtw_Click);
            //
            // textBoxBtwPath
            //
            this.textBoxBtwPath.Location = new System.Drawing.Point(80, 25);
            this.textBoxBtwPath.Name = "textBoxBtwPath";
            this.textBoxBtwPath.ReadOnly = true;
            this.textBoxBtwPath.Size = new System.Drawing.Size(220, 22);
            this.textBoxBtwPath.TabIndex = 1;
            //
            // labelBtwFile
            //
            this.labelBtwFile.AutoSize = true;
            this.labelBtwFile.Location = new System.Drawing.Point(15, 28);
            this.labelBtwFile.Name = "labelBtwFile";
            this.labelBtwFile.Size = new System.Drawing.Size(59, 12);
            this.labelBtwFile.TabIndex = 0;
            this.labelBtwFile.Text = "BTW檔案:";
            //
            // groupBoxInput
            //
            this.groupBoxInput.Controls.Add(this.comboBoxDocCategory);
            this.groupBoxInput.Controls.Add(this.labelDocCategory);
            this.groupBoxInput.Controls.Add(this.buttonGetData);
            this.groupBoxInput.Controls.Add(this.textBoxDocItem);
            this.groupBoxInput.Controls.Add(this.textBoxDocNumber);
            this.groupBoxInput.Controls.Add(this.textBoxDocType);
            this.groupBoxInput.Controls.Add(this.labelDocItem);
            this.groupBoxInput.Controls.Add(this.labelDocNumber);
            this.groupBoxInput.Controls.Add(this.labelDocType);
            this.groupBoxInput.Location = new System.Drawing.Point(12, 370);
            this.groupBoxInput.Name = "groupBoxInput";
            this.groupBoxInput.Size = new System.Drawing.Size(380, 120);
            this.groupBoxInput.TabIndex = 1;
            this.groupBoxInput.TabStop = false;
            this.groupBoxInput.Text = "輸入欄位";
            //
            // buttonGetData
            //
            this.buttonGetData.Location = new System.Drawing.Point(290, 85);
            this.buttonGetData.Name = "buttonGetData";
            this.buttonGetData.Size = new System.Drawing.Size(75, 23);
            this.buttonGetData.TabIndex = 6;
            this.buttonGetData.Text = "獲取資料";
            this.buttonGetData.UseVisualStyleBackColor = true;
            this.buttonGetData.Click += new System.EventHandler(this.buttonGetData_Click);
            //
            // textBoxDocItem
            //
            this.textBoxDocItem.Location = new System.Drawing.Point(80, 85);
            this.textBoxDocItem.Name = "textBoxDocItem";
            this.textBoxDocItem.Size = new System.Drawing.Size(200, 22);
            this.textBoxDocItem.TabIndex = 5;
            //
            // textBoxDocNumber
            //
            this.textBoxDocNumber.Location = new System.Drawing.Point(80, 55);
            this.textBoxDocNumber.Name = "textBoxDocNumber";
            this.textBoxDocNumber.Size = new System.Drawing.Size(200, 22);
            this.textBoxDocNumber.TabIndex = 4;
            //
            // textBoxDocType
            //
            this.textBoxDocType.Location = new System.Drawing.Point(80, 25);
            this.textBoxDocType.Name = "textBoxDocType";
            this.textBoxDocType.Size = new System.Drawing.Size(120, 22);
            this.textBoxDocType.TabIndex = 3;
            //
            // comboBoxDocCategory
            //
            this.comboBoxDocCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxDocCategory.FormattingEnabled = true;
            this.comboBoxDocCategory.Items.AddRange(new object[] {
            "一般進貨",
            "托外進貨"});
            this.comboBoxDocCategory.Location = new System.Drawing.Point(280, 25);
            this.comboBoxDocCategory.Name = "comboBoxDocCategory";
            this.comboBoxDocCategory.Size = new System.Drawing.Size(90, 20);
            this.comboBoxDocCategory.TabIndex = 8;
            this.comboBoxDocCategory.SelectedIndex = 0;
            //
            // labelDocCategory
            //
            this.labelDocCategory.AutoSize = true;
            this.labelDocCategory.Location = new System.Drawing.Point(210, 28);
            this.labelDocCategory.Name = "labelDocCategory";
            this.labelDocCategory.Size = new System.Drawing.Size(59, 12);
            this.labelDocCategory.TabIndex = 7;
            this.labelDocCategory.Text = "單據類型:";
            //
            // labelDocItem
            //
            this.labelDocItem.AutoSize = true;
            this.labelDocItem.Location = new System.Drawing.Point(15, 88);
            this.labelDocItem.Name = "labelDocItem";
            this.labelDocItem.Size = new System.Drawing.Size(59, 12);
            this.labelDocItem.TabIndex = 2;
            this.labelDocItem.Text = "進貨項次:";
            //
            // labelDocNumber
            //
            this.labelDocNumber.AutoSize = true;
            this.labelDocNumber.Location = new System.Drawing.Point(15, 58);
            this.labelDocNumber.Name = "labelDocNumber";
            this.labelDocNumber.Size = new System.Drawing.Size(59, 12);
            this.labelDocNumber.TabIndex = 1;
            this.labelDocNumber.Text = "進貨單號:";
            //
            // labelDocType
            //
            this.labelDocType.AutoSize = true;
            this.labelDocType.Location = new System.Drawing.Point(15, 28);
            this.labelDocType.Name = "labelDocType";
            this.labelDocType.Size = new System.Drawing.Size(59, 12);
            this.labelDocType.TabIndex = 0;
            this.labelDocType.Text = "進貨單別:";
            //
            // groupBoxApiData
            //
            this.groupBoxApiData.Controls.Add(this.richTextBoxApiResult);
            this.groupBoxApiData.Location = new System.Drawing.Point(400, 370);
            this.groupBoxApiData.Name = "groupBoxApiData";
            this.groupBoxApiData.Size = new System.Drawing.Size(788, 120);
            this.groupBoxApiData.TabIndex = 2;
            this.groupBoxApiData.TabStop = false;
            this.groupBoxApiData.Text = "API 回傳資料";
            //
            // richTextBoxApiResult
            //
            this.richTextBoxApiResult.Location = new System.Drawing.Point(10, 20);
            this.richTextBoxApiResult.Name = "richTextBoxApiResult";
            this.richTextBoxApiResult.ReadOnly = true;
            this.richTextBoxApiResult.Size = new System.Drawing.Size(770, 90);
            this.richTextBoxApiResult.TabIndex = 0;
            this.richTextBoxApiResult.Text = "";
            //
            // groupBoxPrintSettings
            //
            this.groupBoxPrintSettings.Controls.Add(this.labelProductNameLength);
            this.groupBoxPrintSettings.Controls.Add(this.textBoxProductName);
            this.groupBoxPrintSettings.Controls.Add(this.labelProductNameTitle);
            this.groupBoxPrintSettings.Controls.Add(this.buttonPrint);
            this.groupBoxPrintSettings.Controls.Add(this.buttonUpdatePreview);
            this.groupBoxPrintSettings.Controls.Add(this.textBoxFwVer);
            this.groupBoxPrintSettings.Controls.Add(this.textBoxHwVer);
            this.groupBoxPrintSettings.Controls.Add(this.textBoxDc);
            this.groupBoxPrintSettings.Controls.Add(this.textBoxQuantity);
            this.groupBoxPrintSettings.Controls.Add(this.labelFwVer);
            this.groupBoxPrintSettings.Controls.Add(this.labelHwVer);
            this.groupBoxPrintSettings.Controls.Add(this.labelDc);
            this.groupBoxPrintSettings.Controls.Add(this.labelQuantity);
            this.groupBoxPrintSettings.Location = new System.Drawing.Point(12, 500);
            this.groupBoxPrintSettings.Name = "groupBoxPrintSettings";
            this.groupBoxPrintSettings.Size = new System.Drawing.Size(1176, 150);
            this.groupBoxPrintSettings.TabIndex = 3;
            this.groupBoxPrintSettings.TabStop = false;
            this.groupBoxPrintSettings.Text = "列印設定";
            //
            // buttonPrint
            //
            this.buttonPrint.Location = new System.Drawing.Point(1060, 80);
            this.buttonPrint.Name = "buttonPrint";
            this.buttonPrint.Size = new System.Drawing.Size(100, 30);
            this.buttonPrint.TabIndex = 8;
            this.buttonPrint.Text = "列印標籤";
            this.buttonPrint.UseVisualStyleBackColor = true;
            this.buttonPrint.Click += new System.EventHandler(this.buttonPrint_Click);
            //
            // buttonUpdatePreview
            //
            this.buttonUpdatePreview.Location = new System.Drawing.Point(730, 25);
            this.buttonUpdatePreview.Name = "buttonUpdatePreview";
            this.buttonUpdatePreview.Size = new System.Drawing.Size(100, 50);
            this.buttonUpdatePreview.TabIndex = 9;
            this.buttonUpdatePreview.Text = "更新預覽";
            this.buttonUpdatePreview.UseVisualStyleBackColor = true;
            this.buttonUpdatePreview.Click += new System.EventHandler(this.buttonUpdatePreview_Click);
            //
            // textBoxFwVer
            //
            this.textBoxFwVer.Location = new System.Drawing.Point(450, 55);
            this.textBoxFwVer.Name = "textBoxFwVer";
            this.textBoxFwVer.Size = new System.Drawing.Size(150, 22);
            this.textBoxFwVer.TabIndex = 7;
            //
            // textBoxHwVer
            //
            this.textBoxHwVer.Location = new System.Drawing.Point(450, 25);
            this.textBoxHwVer.Name = "textBoxHwVer";
            this.textBoxHwVer.Size = new System.Drawing.Size(150, 22);
            this.textBoxHwVer.TabIndex = 6;
            //
            // textBoxDc
            //
            this.textBoxDc.Location = new System.Drawing.Point(80, 55);
            this.textBoxDc.Name = "textBoxDc";
            this.textBoxDc.Size = new System.Drawing.Size(150, 22);
            this.textBoxDc.TabIndex = 5;
            //
            // textBoxQuantity
            //
            this.textBoxQuantity.Location = new System.Drawing.Point(80, 25);
            this.textBoxQuantity.Name = "textBoxQuantity";
            this.textBoxQuantity.Size = new System.Drawing.Size(150, 22);
            this.textBoxQuantity.TabIndex = 4;
            //
            // labelFwVer
            //
            this.labelFwVer.AutoSize = true;
            this.labelFwVer.Location = new System.Drawing.Point(390, 58);
            this.labelFwVer.Name = "labelFwVer";
            this.labelFwVer.Size = new System.Drawing.Size(48, 12);
            this.labelFwVer.TabIndex = 3;
            this.labelFwVer.Text = "FW Ver:";
            //
            // labelHwVer
            //
            this.labelHwVer.AutoSize = true;
            this.labelHwVer.Location = new System.Drawing.Point(390, 28);
            this.labelHwVer.Name = "labelHwVer";
            this.labelHwVer.Size = new System.Drawing.Size(50, 12);
            this.labelHwVer.TabIndex = 2;
            this.labelHwVer.Text = "HW Ver:";
            //
            // labelDc
            //
            this.labelDc.AutoSize = true;
            this.labelDc.Location = new System.Drawing.Point(15, 58);
            this.labelDc.Name = "labelDc";
            this.labelDc.Size = new System.Drawing.Size(28, 12);
            this.labelDc.TabIndex = 1;
            this.labelDc.Text = "D/C:";
            //
            // labelQuantity
            //
            this.labelQuantity.AutoSize = true;
            this.labelQuantity.Location = new System.Drawing.Point(15, 28);
            this.labelQuantity.Name = "labelQuantity";
            this.labelQuantity.Size = new System.Drawing.Size(35, 12);
            this.labelQuantity.TabIndex = 0;
            this.labelQuantity.Text = "數量:";
            //
            // labelProductNameTitle
            //
            this.labelProductNameTitle.AutoSize = true;
            this.labelProductNameTitle.Location = new System.Drawing.Point(20, 95);
            this.labelProductNameTitle.Name = "labelProductNameTitle";
            this.labelProductNameTitle.Size = new System.Drawing.Size(89, 12);
            this.labelProductNameTitle.TabIndex = 10;
            this.labelProductNameTitle.Text = "品名（可編輯）：";
            //
            // textBoxProductName
            //
            this.textBoxProductName.Location = new System.Drawing.Point(115, 92);
            this.textBoxProductName.Multiline = true;
            this.textBoxProductName.Name = "textBoxProductName";
            this.textBoxProductName.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.textBoxProductName.Size = new System.Drawing.Size(600, 40);
            this.textBoxProductName.TabIndex = 11;
            this.textBoxProductName.TextChanged += new System.EventHandler(this.textBoxProductName_TextChanged);
            //
            // labelProductNameLength
            //
            this.labelProductNameLength.AutoSize = true;
            this.labelProductNameLength.ForeColor = System.Drawing.Color.Gray;
            this.labelProductNameLength.Location = new System.Drawing.Point(730, 95);
            this.labelProductNameLength.Name = "labelProductNameLength";
            this.labelProductNameLength.Size = new System.Drawing.Size(53, 12);
            this.labelProductNameLength.TabIndex = 12;
            this.labelProductNameLength.Text = "長度: 0";
            //
            // openFileDialogBtw
            //
            this.openFileDialogBtw.Filter = "BarTender Files (*.btw)|*.btw|All Files (*.*)|*.*";
            this.openFileDialogBtw.Title = "選擇 BTW 檔案";
            //
            // Form1
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 700);
            this.Controls.Add(this.groupBoxPrintSettings);
            this.Controls.Add(this.groupBoxApiData);
            this.Controls.Add(this.groupBoxInput);
            this.Controls.Add(this.groupBoxBtwFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "標籤列印管理程式";
            this.groupBoxBtwFile.ResumeLayout(false);
            this.groupBoxBtwFile.PerformLayout();
            this.groupBoxInput.ResumeLayout(false);
            this.groupBoxInput.PerformLayout();
            this.groupBoxApiData.ResumeLayout(false);
            this.groupBoxPrintSettings.ResumeLayout(false);
            this.groupBoxPrintSettings.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxBtwFile;
        private System.Windows.Forms.PictureBox pictureBoxPreview;
        private System.Windows.Forms.Button buttonBrowseBtw;
        private System.Windows.Forms.TextBox textBoxBtwPath;
        private System.Windows.Forms.Label labelBtwFile;
        private System.Windows.Forms.GroupBox groupBoxInput;
        private System.Windows.Forms.Button buttonGetData;
        private System.Windows.Forms.TextBox textBoxDocItem;
        private System.Windows.Forms.TextBox textBoxDocNumber;
        private System.Windows.Forms.TextBox textBoxDocType;
        private System.Windows.Forms.Label labelDocItem;
        private System.Windows.Forms.Label labelDocNumber;
        private System.Windows.Forms.Label labelDocType;
        private System.Windows.Forms.GroupBox groupBoxApiData;
        private System.Windows.Forms.RichTextBox richTextBoxApiResult;
        private System.Windows.Forms.GroupBox groupBoxPrintSettings;
        private System.Windows.Forms.Button buttonPrint;
        private System.Windows.Forms.Button buttonUpdatePreview;
        private System.Windows.Forms.TextBox textBoxFwVer;
        private System.Windows.Forms.TextBox textBoxHwVer;
        private System.Windows.Forms.TextBox textBoxDc;
        private System.Windows.Forms.TextBox textBoxQuantity;
        private System.Windows.Forms.Label labelFwVer;
        private System.Windows.Forms.Label labelHwVer;
        private System.Windows.Forms.Label labelDc;
        private System.Windows.Forms.Label labelQuantity;
        private System.Windows.Forms.Label labelProductNameTitle;
        private System.Windows.Forms.TextBox textBoxProductName;
        private System.Windows.Forms.Label labelProductNameLength;
        private System.Windows.Forms.OpenFileDialog openFileDialogBtw;
        private System.Windows.Forms.ComboBox comboBoxDocCategory;
        private System.Windows.Forms.Label labelDocCategory;
    }
}

