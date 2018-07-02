namespace Secretarium.Excel
{
    partial class LoadX509
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.x509LoadErrorLabel = new System.Windows.Forms.Label();
            this.x509LoadImg = new System.Windows.Forms.PictureBox();
            this.x509BtnLoad = new System.Windows.Forms.Button();
            this.x509PasswordInput = new System.Windows.Forms.TextBox();
            this.x509PasswordLabel = new System.Windows.Forms.Label();
            this.x509PathInput = new System.Windows.Forms.TextBox();
            this.x509Browse = new System.Windows.Forms.Button();
            this.x509BrowseLabel = new System.Windows.Forms.Label();
            this.x509OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.x509LoadImg)).BeginInit();
            this.SuspendLayout();
            // 
            // x509LoadErrorLabel
            // 
            this.x509LoadErrorLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.x509LoadErrorLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509LoadErrorLabel.ForeColor = System.Drawing.Color.Red;
            this.x509LoadErrorLabel.Location = new System.Drawing.Point(12, 460);
            this.x509LoadErrorLabel.Name = "x509LoadErrorLabel";
            this.x509LoadErrorLabel.Size = new System.Drawing.Size(405, 65);
            this.x509LoadErrorLabel.TabIndex = 0;
            // 
            // x509LoadImg
            // 
            this.x509LoadImg.ErrorImage = null;
            this.x509LoadImg.InitialImage = null;
            this.x509LoadImg.Location = new System.Drawing.Point(233, 385);
            this.x509LoadImg.Name = "x509LoadImg";
            this.x509LoadImg.Size = new System.Drawing.Size(54, 54);
            this.x509LoadImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.x509LoadImg.TabIndex = 6;
            this.x509LoadImg.TabStop = false;
            // 
            // x509BtnLoad
            // 
            this.x509BtnLoad.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(170)))), ((int)(((byte)(240)))));
            this.x509BtnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.x509BtnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509BtnLoad.ForeColor = System.Drawing.Color.White;
            this.x509BtnLoad.Location = new System.Drawing.Point(17, 385);
            this.x509BtnLoad.Name = "x509BtnLoad";
            this.x509BtnLoad.Size = new System.Drawing.Size(188, 54);
            this.x509BtnLoad.TabIndex = 4;
            this.x509BtnLoad.Text = "Load";
            this.x509BtnLoad.UseVisualStyleBackColor = false;
            this.x509BtnLoad.Click += new System.EventHandler(this.x509BtnLoad_Click);
            // 
            // x509PasswordInput
            // 
            this.x509PasswordInput.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509PasswordInput.Location = new System.Drawing.Point(16, 289);
            this.x509PasswordInput.MinimumSize = new System.Drawing.Size(400, 54);
            this.x509PasswordInput.Name = "x509PasswordInput";
            this.x509PasswordInput.PasswordChar = '●';
            this.x509PasswordInput.Size = new System.Drawing.Size(400, 54);
            this.x509PasswordInput.TabIndex = 3;
            this.x509PasswordInput.UseSystemPasswordChar = true;
            this.x509PasswordInput.AutoSize = false;
            // 
            // x509PasswordLabel
            // 
            this.x509PasswordLabel.AutoSize = true;
            this.x509PasswordLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509PasswordLabel.Location = new System.Drawing.Point(12, 248);
            this.x509PasswordLabel.Name = "x509PasswordLabel";
            this.x509PasswordLabel.Size = new System.Drawing.Size(120, 29);
            this.x509PasswordLabel.TabIndex = 0;
            this.x509PasswordLabel.Text = "Password";
            // 
            // x509PathInput
            // 
            this.x509PathInput.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509PathInput.Location = new System.Drawing.Point(16, 90);
            this.x509PathInput.MinimumSize = new System.Drawing.Size(340, 54);
            this.x509PathInput.Name = "x509PathInput";
            this.x509PathInput.Size = new System.Drawing.Size(400, 54);
            this.x509PathInput.TabIndex = 1;
            this.x509PathInput.AutoSize = false;
            // 
            // x509Browse
            // 
            this.x509Browse.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(170)))), ((int)(((byte)(240)))));
            this.x509Browse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.x509Browse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509Browse.ForeColor = System.Drawing.Color.White;
            this.x509Browse.Location = new System.Drawing.Point(17, 150);
            this.x509Browse.Name = "x509Browse";
            this.x509Browse.Size = new System.Drawing.Size(128, 54);
            this.x509Browse.TabIndex = 2;
            this.x509Browse.Text = "Browse";
            this.x509Browse.UseVisualStyleBackColor = false;
            this.x509Browse.Click += new System.EventHandler(this.x509Browse_Click);
            // 
            // x509BrowseLabel
            // 
            this.x509BrowseLabel.AutoSize = true;
            this.x509BrowseLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.x509BrowseLabel.Location = new System.Drawing.Point(12, 49);
            this.x509BrowseLabel.Name = "x509BrowseLabel";
            this.x509BrowseLabel.Size = new System.Drawing.Size(321, 29);
            this.x509BrowseLabel.TabIndex = 0;
            this.x509BrowseLabel.Text = "Load from certificate (pfx file)";
            // 
            // x509OpenFileDialog
            // 
            this.x509OpenFileDialog.DefaultExt = "pfx";
            this.x509OpenFileDialog.FileName = "YourPfxCertificate";
            this.x509OpenFileDialog.Filter = "pfx files|*.pfx|All files|*.*";
            this.x509OpenFileDialog.Title = "Load your X509 pfx certificate";
            // 
            // LoadX509
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.x509LoadErrorLabel);
            this.Controls.Add(this.x509LoadImg);
            this.Controls.Add(this.x509BtnLoad);
            this.Controls.Add(this.x509BrowseLabel);
            this.Controls.Add(this.x509PasswordInput);
            this.Controls.Add(this.x509Browse);
            this.Controls.Add(this.x509PasswordLabel);
            this.Controls.Add(this.x509PathInput);
            this.MinimumSize = new System.Drawing.Size(120, 300);
            this.Name = "LoadX509";
            this.Size = new System.Drawing.Size(520, 686);
            ((System.ComponentModel.ISupportInitialize)(this.x509LoadImg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label x509BrowseLabel;
        private System.Windows.Forms.TextBox x509PathInput;
        private System.Windows.Forms.Button x509Browse;
        private System.Windows.Forms.TextBox x509PasswordInput;
        private System.Windows.Forms.Label x509PasswordLabel;
        private System.Windows.Forms.PictureBox x509LoadImg;
        private System.Windows.Forms.Button x509BtnLoad;
        private System.Windows.Forms.OpenFileDialog x509OpenFileDialog;
        private System.Windows.Forms.Label x509LoadErrorLabel;
    }
}
