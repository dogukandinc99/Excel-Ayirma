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
            fileselectbtn = new Button();
            adresstxt = new TextBox();
            cellvaluebtn = new Button();
            cellvaluetxt = new TextBox();
            dataGridView1 = new DataGridView();
            listbtn = new Button();
            progressBar1 = new ProgressBar();
            listBox1 = new ListBox();
            excelclosebtn = new Button();
            saveadressfoldertxt = new TextBox();
            saveselectedfolderbtn = new Button();
            saveexcelbtn = new Button();
            numericUpDown1 = new NumericUpDown();
            testbtn = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown1).BeginInit();
            SuspendLayout();
            // 
            // fileselectbtn
            // 
            fileselectbtn.Location = new Point(566, 9);
            fileselectbtn.Name = "fileselectbtn";
            fileselectbtn.Size = new Size(150, 23);
            fileselectbtn.TabIndex = 0;
            fileselectbtn.Text = "Dosya Seç";
            fileselectbtn.UseVisualStyleBackColor = true;
            fileselectbtn.Click += fileselectbtn_Click;
            // 
            // adresstxt
            // 
            adresstxt.Enabled = false;
            adresstxt.Location = new Point(12, 9);
            adresstxt.Name = "adresstxt";
            adresstxt.Size = new Size(525, 23);
            adresstxt.TabIndex = 1;
            // 
            // cellvaluebtn
            // 
            cellvaluebtn.Enabled = false;
            cellvaluebtn.Location = new Point(566, 38);
            cellvaluebtn.Name = "cellvaluebtn";
            cellvaluebtn.Size = new Size(150, 23);
            cellvaluebtn.TabIndex = 2;
            cellvaluebtn.Text = "Hücreyi getir";
            cellvaluebtn.UseVisualStyleBackColor = true;
            cellvaluebtn.Click += cellvaluebtn_Click;
            // 
            // cellvaluetxt
            // 
            cellvaluetxt.Enabled = false;
            cellvaluetxt.Location = new Point(12, 38);
            cellvaluetxt.Name = "cellvaluetxt";
            cellvaluetxt.Size = new Size(525, 23);
            cellvaluetxt.TabIndex = 3;
            cellvaluetxt.Text = "1-15";
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(12, 94);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(525, 247);
            dataGridView1.TabIndex = 4;
            // 
            // listbtn
            // 
            listbtn.Enabled = false;
            listbtn.Location = new Point(722, 249);
            listbtn.Name = "listbtn";
            listbtn.Size = new Size(65, 92);
            listbtn.TabIndex = 5;
            listbtn.Text = "Listele";
            listbtn.UseVisualStyleBackColor = true;
            listbtn.Click += listbtn_Click;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(12, 357);
            progressBar1.Maximum = 10000;
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(705, 23);
            progressBar1.TabIndex = 6;
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 15;
            listBox1.Location = new Point(566, 95);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(150, 244);
            listBox1.TabIndex = 7;
            // 
            // excelclosebtn
            // 
            excelclosebtn.Enabled = false;
            excelclosebtn.Location = new Point(722, 8);
            excelclosebtn.Name = "excelclosebtn";
            excelclosebtn.Size = new Size(65, 79);
            excelclosebtn.TabIndex = 8;
            excelclosebtn.Text = "Exceli Kapat";
            excelclosebtn.UseVisualStyleBackColor = true;
            excelclosebtn.Click += excelclosebtn_Click;
            // 
            // saveadressfoldertxt
            // 
            saveadressfoldertxt.Enabled = false;
            saveadressfoldertxt.Location = new Point(12, 67);
            saveadressfoldertxt.Name = "saveadressfoldertxt";
            saveadressfoldertxt.Size = new Size(525, 23);
            saveadressfoldertxt.TabIndex = 9;
            // 
            // saveselectedfolderbtn
            // 
            saveselectedfolderbtn.Location = new Point(566, 68);
            saveselectedfolderbtn.Name = "saveselectedfolderbtn";
            saveselectedfolderbtn.Size = new Size(150, 23);
            saveselectedfolderbtn.TabIndex = 10;
            saveselectedfolderbtn.Text = "Kaydedilecek Dizin";
            saveselectedfolderbtn.UseVisualStyleBackColor = true;
            saveselectedfolderbtn.Click += saveselectedfolderbtn_Click;
            // 
            // saveexcelbtn
            // 
            saveexcelbtn.Location = new Point(722, 95);
            saveexcelbtn.Name = "saveexcelbtn";
            saveexcelbtn.Size = new Size(65, 95);
            saveexcelbtn.TabIndex = 11;
            saveexcelbtn.Text = "Kaydet";
            saveexcelbtn.UseVisualStyleBackColor = true;
            saveexcelbtn.Click += saveexcelbtn_Click;
            // 
            // numericUpDown1
            // 
            numericUpDown1.Location = new Point(725, 357);
            numericUpDown1.Name = "numericUpDown1";
            numericUpDown1.Size = new Size(62, 23);
            numericUpDown1.TabIndex = 12;
            numericUpDown1.Value = new decimal(new int[] { 3, 0, 0, 0 });
            // 
            // testbtn
            // 
            testbtn.Location = new Point(722, 196);
            testbtn.Name = "testbtn";
            testbtn.Size = new Size(65, 44);
            testbtn.TabIndex = 13;
            testbtn.Text = "Test";
            testbtn.UseVisualStyleBackColor = true;
            testbtn.Click += testbtn_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(798, 395);
            Controls.Add(testbtn);
            Controls.Add(numericUpDown1);
            Controls.Add(saveexcelbtn);
            Controls.Add(saveselectedfolderbtn);
            Controls.Add(saveadressfoldertxt);
            Controls.Add(excelclosebtn);
            Controls.Add(listBox1);
            Controls.Add(progressBar1);
            Controls.Add(listbtn);
            Controls.Add(dataGridView1);
            Controls.Add(cellvaluetxt);
            Controls.Add(cellvaluebtn);
            Controls.Add(adresstxt);
            Controls.Add(fileselectbtn);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown1).EndInit();
            ResumeLayout(false);
            PerformLayout();
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
        private Button testbtn;
    }
}