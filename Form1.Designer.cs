namespace GetCodeTool
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.dataGridViewOTP = new System.Windows.Forms.DataGridView();
            this.btnGetCode = new System.Windows.Forms.Button();
            this.loadingPictureBox = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewOTP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.loadingPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // txtEmail
            // 
            this.txtEmail.Location = new System.Drawing.Point(12, 12);
            this.txtEmail.Multiline = true;
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(375, 538);
            this.txtEmail.TabIndex = 1;
            this.txtEmail.Enter += new System.EventHandler(this.txtEmail_Enter);
            this.txtEmail.Leave += new System.EventHandler(this.txtEmail_Leave);
            // 
            // dataGridViewOTP
            // 
            this.dataGridViewOTP.AllowUserToAddRows = false;
            this.dataGridViewOTP.AllowUserToDeleteRows = false;
            this.dataGridViewOTP.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridViewOTP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewOTP.Location = new System.Drawing.Point(393, 12);
            this.dataGridViewOTP.Name = "dataGridViewOTP";
            this.dataGridViewOTP.ReadOnly = true;
            this.dataGridViewOTP.Size = new System.Drawing.Size(778, 538);
            this.dataGridViewOTP.TabIndex = 2;
            // 
            // btnGetCode
            // 
            this.btnGetCode.Location = new System.Drawing.Point(1059, 625);
            this.btnGetCode.Name = "btnGetCode";
            this.btnGetCode.Size = new System.Drawing.Size(112, 37);
            this.btnGetCode.TabIndex = 3;
            this.btnGetCode.Text = "button1";
            this.btnGetCode.UseVisualStyleBackColor = true;
            this.btnGetCode.Click += new System.EventHandler(this.btnGetCode_Click);
            // 
            // loadingPictureBox
            // 
            this.loadingPictureBox.Location = new System.Drawing.Point(366, 556);
            this.loadingPictureBox.Name = "loadingPictureBox";
            this.loadingPictureBox.Size = new System.Drawing.Size(100, 89);
            this.loadingPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.loadingPictureBox.TabIndex = 4;
            this.loadingPictureBox.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1183, 674);
            this.Controls.Add(this.loadingPictureBox);
            this.Controls.Add(this.btnGetCode);
            this.Controls.Add(this.dataGridViewOTP);
            this.Controls.Add(this.txtEmail);
            this.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewOTP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.loadingPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.DataGridView dataGridViewOTP;
        private System.Windows.Forms.Button btnGetCode;
        private System.Windows.Forms.PictureBox loadingPictureBox;
    }
}

