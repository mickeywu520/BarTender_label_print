using System;
using System.Drawing;
using System.Windows.Forms;

namespace LabelPrintManager
{
    /// <summary>
    /// 進度對話框
    /// </summary>
    public partial class ProgressForm : Form
    {
        private Label labelMessage;
        private ProgressBar progressBar;
        private Button buttonCancel;
        private Timer animationTimer;
        private int animationStep = 0;

        public bool IsCancelled { get; private set; } = false;

        public ProgressForm(string message)
        {
            InitializeComponent();
            SetMessage(message);
            StartAnimation();
        }

        private void InitializeComponent()
        {
            this.labelMessage = new Label();
            this.progressBar = new ProgressBar();
            this.buttonCancel = new Button();
            this.animationTimer = new Timer();
            this.SuspendLayout();

            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = false;
            this.labelMessage.Location = new Point(20, 20);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new Size(300, 40);
            this.labelMessage.TabIndex = 0;
            this.labelMessage.Text = "正在處理...";
            this.labelMessage.TextAlign = ContentAlignment.MiddleCenter;

            // 
            // progressBar
            // 
            this.progressBar.Location = new Point(20, 70);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new Size(300, 23);
            this.progressBar.Style = ProgressBarStyle.Marquee;
            this.progressBar.MarqueeAnimationSpeed = 50;
            this.progressBar.TabIndex = 1;

            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new Point(135, 110);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new Size(75, 30);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "取消";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new EventHandler(this.buttonCancel_Click);

            // 
            // animationTimer
            // 
            this.animationTimer.Interval = 500;
            this.animationTimer.Tick += new EventHandler(this.animationTimer_Tick);

            // 
            // ProgressForm
            // 
            this.AutoScaleDimensions = new SizeF(6F, 12F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(340, 160);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.labelMessage);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "列印中";
            this.TopMost = true;
            this.ResumeLayout(false);
        }

        public void SetMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(SetMessage), message);
                return;
            }

            this.labelMessage.Text = message;
        }

        private void StartAnimation()
        {
            animationTimer.Start();
        }

        private void StopAnimation()
        {
            animationTimer.Stop();
        }

        private void animationTimer_Tick(object sender, EventArgs e)
        {
            animationStep = (animationStep + 1) % 4;
            string dots = new string('.', animationStep + 1);
            string baseMessage = labelMessage.Text.TrimEnd('.');
            
            // 移除之前的點
            int lastDotIndex = baseMessage.LastIndexOf('.');
            if (lastDotIndex > 0)
            {
                baseMessage = baseMessage.Substring(0, lastDotIndex);
            }
            
            labelMessage.Text = baseMessage + dots;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            IsCancelled = true;
            StopAnimation();
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            StopAnimation();
            base.OnFormClosing(e);
        }

        /// <summary>
        /// 禁用關閉按鈕（只能通過取消按鈕關閉）
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= 0x200; // CS_NOCLOSE
                return cp;
            }
        }
    }
}
