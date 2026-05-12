namespace ExtractStatementPDF
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
            button1 = new Button();
            richTextBox1 = new RichTextBox();
            button2 = new Button();
            txtARStatement = new TextBox();
            txtRxOffice = new TextBox();
            textBox3 = new TextBox();
            button3 = new Button();
            button4 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(12, 12);
            button1.Name = "button1";
            button1.Size = new Size(123, 23);
            button1.TabIndex = 0;
            button1.Text = "Upload PDF";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(356, 203);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(576, 315);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // button2
            // 
            button2.Location = new Point(12, 41);
            button2.Name = "button2";
            button2.Size = new Size(123, 23);
            button2.TabIndex = 2;
            button2.Text = "Upload CSV";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // txtARStatement
            // 
            txtARStatement.Location = new Point(160, 13);
            txtARStatement.Name = "txtARStatement";
            txtARStatement.Size = new Size(100, 23);
            txtARStatement.TabIndex = 3;
            // 
            // txtRxOffice
            // 
            txtRxOffice.Location = new Point(266, 13);
            txtRxOffice.Name = "txtRxOffice";
            txtRxOffice.Size = new Size(100, 23);
            txtRxOffice.TabIndex = 4;
            // 
            // textBox3
            // 
            textBox3.Location = new Point(372, 13);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(100, 23);
            textBox3.TabIndex = 5;
            // 
            // button3
            // 
            button3.Location = new Point(12, 70);
            button3.Name = "button3";
            button3.Size = new Size(123, 23);
            button3.TabIndex = 6;
            button3.Text = "Download Excel";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(12, 177);
            button4.Name = "button4";
            button4.Size = new Size(123, 23);
            button4.TabIndex = 7;
            button4.Text = "Scan Folder";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(944, 530);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(textBox3);
            Controls.Add(txtRxOffice);
            Controls.Add(txtARStatement);
            Controls.Add(button2);
            Controls.Add(richTextBox1);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private RichTextBox richTextBox1;
        private Button button2;
        private TextBox txtARStatement;
        private TextBox txtRxOffice;
        private TextBox textBox3;
        private Button button3;
        private Button button4;
    }
}
