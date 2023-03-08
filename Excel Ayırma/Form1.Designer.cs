namespace Excel_Ayırma
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.fileselectbtn = new System.Windows.Forms.Button();
            this.adresstxt = new System.Windows.Forms.TextBox();
            this.cellvaluebtn = new System.Windows.Forms.Button();
            this.cellvaluetxt = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.listbtn = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.excelclosebtn = new System.Windows.Forms.Button();
            this.saveadressfoldertxt = new System.Windows.Forms.TextBox();
            this.saveselectedfolderbtn = new System.Windows.Forms.Button();
            this.saveexcelbtn = new System.Windows.Forms.Button();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // fileselectbtn
            // 
            this.fileselectbtn.Location = new System.Drawing.Point(420, 11);
            this.fileselectbtn.Name = "fileselectbtn";
            this.fileselectbtn.Size = new System.Drawing.Size(150, 23);
            this.fileselectbtn.TabIndex = 0;
            this.fileselectbtn.Text = "Dosya Seç";
            this.fileselectbtn.UseVisualStyleBackColor = true;
            this.fileselectbtn.Click += new System.EventHandler(this.fileselectbtn_Click);
            // 
            // adresstxt
            // 
            this.adresstxt.Enabled = false;
            this.adresstxt.Location = new System.Drawing.Point(44, 12);
            this.adresstxt.Name = "adresstxt";
            this.adresstxt.Size = new System.Drawing.Size(346, 23);
            this.adresstxt.TabIndex = 1;
            // 
            // cellvaluebtn
            // 
            this.cellvaluebtn.Enabled = false;
            this.cellvaluebtn.Location = new System.Drawing.Point(420, 40);
            this.cellvaluebtn.Name = "cellvaluebtn";
            this.cellvaluebtn.Size = new System.Drawing.Size(150, 23);
            this.cellvaluebtn.TabIndex = 2;
            this.cellvaluebtn.Text = "Hücreyi getir";
            this.cellvaluebtn.UseVisualStyleBackColor = true;
            this.cellvaluebtn.Click += new System.EventHandler(this.cellvaluebtn_Click);
            // 
            // cellvaluetxt
            // 
            this.cellvaluetxt.Enabled = false;
            this.cellvaluetxt.Location = new System.Drawing.Point(44, 41);
            this.cellvaluetxt.Name = "cellvaluetxt";
            this.cellvaluetxt.Size = new System.Drawing.Size(346, 23);
            this.cellvaluetxt.TabIndex = 3;
            this.cellvaluetxt.Text = "1-15";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(44, 97);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(346, 247);
            this.dataGridView1.TabIndex = 4;
            // 
            // listbtn
            // 
            this.listbtn.Enabled = false;
            this.listbtn.Location = new System.Drawing.Point(576, 97);
            this.listbtn.Name = "listbtn";
            this.listbtn.Size = new System.Drawing.Size(65, 124);
            this.listbtn.TabIndex = 5;
            this.listbtn.Text = "Listele";
            this.listbtn.UseVisualStyleBackColor = true;
            this.listbtn.Click += new System.EventHandler(this.listbtn_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(44, 360);
            this.progressBar1.Maximum = 10000;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(526, 23);
            this.progressBar1.TabIndex = 6;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 15;
            this.listBox1.Location = new System.Drawing.Point(420, 97);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(150, 244);
            this.listBox1.TabIndex = 7;
            // 
            // excelclosebtn
            // 
            this.excelclosebtn.Enabled = false;
            this.excelclosebtn.Location = new System.Drawing.Point(576, 11);
            this.excelclosebtn.Name = "excelclosebtn";
            this.excelclosebtn.Size = new System.Drawing.Size(65, 79);
            this.excelclosebtn.TabIndex = 8;
            this.excelclosebtn.Text = "Exceli Kapat";
            this.excelclosebtn.UseVisualStyleBackColor = true;
            this.excelclosebtn.Click += new System.EventHandler(this.excelclosebtn_Click);
            // 
            // saveadressfoldertxt
            // 
            this.saveadressfoldertxt.Enabled = false;
            this.saveadressfoldertxt.Location = new System.Drawing.Point(44, 70);
            this.saveadressfoldertxt.Name = "saveadressfoldertxt";
            this.saveadressfoldertxt.Size = new System.Drawing.Size(346, 23);
            this.saveadressfoldertxt.TabIndex = 9;
            // 
            // saveselectedfolderbtn
            // 
            this.saveselectedfolderbtn.Enabled = false;
            this.saveselectedfolderbtn.Location = new System.Drawing.Point(420, 70);
            this.saveselectedfolderbtn.Name = "saveselectedfolderbtn";
            this.saveselectedfolderbtn.Size = new System.Drawing.Size(150, 23);
            this.saveselectedfolderbtn.TabIndex = 10;
            this.saveselectedfolderbtn.Text = "Kaydedilecek Dizin";
            this.saveselectedfolderbtn.UseVisualStyleBackColor = true;
            this.saveselectedfolderbtn.Click += new System.EventHandler(this.saveselectedfolderbtn_Click);
            // 
            // saveexcelbtn
            // 
            this.saveexcelbtn.Enabled = false;
            this.saveexcelbtn.Location = new System.Drawing.Point(576, 220);
            this.saveexcelbtn.Name = "saveexcelbtn";
            this.saveexcelbtn.Size = new System.Drawing.Size(65, 124);
            this.saveexcelbtn.TabIndex = 11;
            this.saveexcelbtn.Text = "Kaydet";
            this.saveexcelbtn.UseVisualStyleBackColor = true;
            this.saveexcelbtn.Click += new System.EventHandler(this.saveexcelbtn_Click);
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(579, 360);
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(62, 23);
            this.numericUpDown1.TabIndex = 12;
            this.numericUpDown1.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(671, 395);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.saveexcelbtn);
            this.Controls.Add(this.saveselectedfolderbtn);
            this.Controls.Add(this.saveadressfoldertxt);
            this.Controls.Add(this.excelclosebtn);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.listbtn);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.cellvaluetxt);
            this.Controls.Add(this.cellvaluebtn);
            this.Controls.Add(this.adresstxt);
            this.Controls.Add(this.fileselectbtn);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button fileselectbtn;
        private TextBox adresstxt;
        private Button cellvaluebtn;
        private TextBox cellvaluetxt;
        private DataGridView dataGridView1;
        private Button listbtn;
        private ProgressBar progressBar1;
        private ListBox listBox1;
        private Button excelclosebtn;
        private TextBox saveadressfoldertxt;
        private Button saveselectedfolderbtn;
        private Button saveexcelbtn;
        private NumericUpDown numericUpDown1;
    }
}