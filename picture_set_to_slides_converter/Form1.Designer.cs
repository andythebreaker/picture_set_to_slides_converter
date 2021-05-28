namespace picture_set_to_slides_converter
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
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.ofb1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.fbd1 = new System.Windows.Forms.FolderBrowserDialog();
            this.fbd2 = new System.Windows.Forms.FolderBrowserDialog();
            this.obf2 = new System.Windows.Forms.Button();
            this.tb1 = new System.Windows.Forms.TextBox();
            this.tb2 = new System.Windows.Forms.TextBox();
            this.lab1 = new System.Windows.Forms.Label();
            this.lab2 = new System.Windows.Forms.Label();
            this.cb1 = new System.Windows.Forms.ComboBox();
            this.pgx = new System.Windows.Forms.NumericUpDown();
            this.pgy = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pgx)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pgy)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(15, 168);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(134, 45);
            this.button1.TabIndex = 0;
            this.button1.Text = "go!";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(282, 168);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(19, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "...";
            // 
            // ofb1
            // 
            this.ofb1.Location = new System.Drawing.Point(198, 12);
            this.ofb1.Name = "ofb1";
            this.ofb1.Size = new System.Drawing.Size(75, 23);
            this.ofb1.TabIndex = 2;
            this.ofb1.Text = "select";
            this.ofb1.UseVisualStyleBackColor = true;
            this.ofb1.Click += new System.EventHandler(this.ofb1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(195, 168);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 15);
            this.label2.TabIndex = 3;
            this.label2.Text = "save at:";
            // 
            // obf2
            // 
            this.obf2.Location = new System.Drawing.Point(198, 43);
            this.obf2.Name = "obf2";
            this.obf2.Size = new System.Drawing.Size(75, 23);
            this.obf2.TabIndex = 4;
            this.obf2.Text = "select";
            this.obf2.UseVisualStyleBackColor = true;
            this.obf2.Click += new System.EventHandler(this.obf2_Click);
            // 
            // tb1
            // 
            this.tb1.Location = new System.Drawing.Point(295, 13);
            this.tb1.Name = "tb1";
            this.tb1.Size = new System.Drawing.Size(470, 25);
            this.tb1.TabIndex = 5;
            // 
            // tb2
            // 
            this.tb2.Location = new System.Drawing.Point(295, 41);
            this.tb2.Name = "tb2";
            this.tb2.Size = new System.Drawing.Size(470, 25);
            this.tb2.TabIndex = 6;
            // 
            // lab1
            // 
            this.lab1.AutoSize = true;
            this.lab1.Location = new System.Drawing.Point(12, 15);
            this.lab1.Name = "lab1";
            this.lab1.Size = new System.Drawing.Size(137, 15);
            this.lab1.TabIndex = 7;
            this.lab1.Text = "Target storage location";
            // 
            // lab2
            // 
            this.lab2.AutoSize = true;
            this.lab2.Location = new System.Drawing.Point(12, 47);
            this.lab2.Name = "lab2";
            this.lab2.Size = new System.Drawing.Size(150, 15);
            this.lab2.TabIndex = 8;
            this.lab2.Text = "Image collection location";
            this.lab2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // cb1
            // 
            this.cb1.FormattingEnabled = true;
            this.cb1.Location = new System.Drawing.Point(502, 119);
            this.cb1.Name = "cb1";
            this.cb1.Size = new System.Drawing.Size(184, 23);
            this.cb1.TabIndex = 9;
            // 
            // pgx
            // 
            this.pgx.DecimalPlaces = 3;
            this.pgx.Location = new System.Drawing.Point(15, 120);
            this.pgx.Maximum = new decimal(new int[] {
            1200,
            0,
            0,
            0});
            this.pgx.Name = "pgx";
            this.pgx.Size = new System.Drawing.Size(120, 25);
            this.pgx.TabIndex = 10;
            this.pgx.Value = new decimal(new int[] {
            254,
            0,
            0,
            65536});
            // 
            // pgy
            // 
            this.pgy.DecimalPlaces = 3;
            this.pgy.Location = new System.Drawing.Point(269, 120);
            this.pgy.Maximum = new decimal(new int[] {
            1200,
            0,
            0,
            0});
            this.pgy.Name = "pgy";
            this.pgy.Size = new System.Drawing.Size(120, 25);
            this.pgy.TabIndex = 11;
            this.pgy.Value = new decimal(new int[] {
            14288,
            0,
            0,
            196608});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(184, 15);
            this.label3.TabIndex = 12;
            this.label3.Text = "Film horizontal axis centimeter";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(272, 85);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(136, 15);
            this.label4.TabIndex = 13;
            this.label4.Text = "Film height centimeter";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(499, 85);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 15);
            this.label5.TabIndex = 14;
            this.label5.Text = "Aspect ratio";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(195, 198);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(85, 15);
            this.label6.TabIndex = 15;
            this.label6.Text = "picture count:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(309, 198);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(19, 15);
            this.label7.TabIndex = 16;
            this.label7.Text = "...";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(859, 301);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pgy);
            this.Controls.Add(this.pgx);
            this.Controls.Add(this.cb1);
            this.Controls.Add(this.lab2);
            this.Controls.Add(this.lab1);
            this.Controls.Add(this.tb2);
            this.Controls.Add(this.tb1);
            this.Controls.Add(this.obf2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ofb1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Batch picture slideshow production";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pgx)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pgy)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button ofb1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog fbd1;
        private System.Windows.Forms.FolderBrowserDialog fbd2;
        private System.Windows.Forms.Button obf2;
        private System.Windows.Forms.TextBox tb1;
        private System.Windows.Forms.TextBox tb2;
        private System.Windows.Forms.Label lab1;
        private System.Windows.Forms.Label lab2;
        private System.Windows.Forms.ComboBox cb1;
        private System.Windows.Forms.NumericUpDown pgx;
        private System.Windows.Forms.NumericUpDown pgy;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
    }
}

