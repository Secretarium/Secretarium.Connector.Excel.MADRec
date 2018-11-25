namespace Secretarium.Excel
{
    partial class LoadSecKey
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
            this.secKeyLoadErrorLabel = new System.Windows.Forms.Label();
            this.secKeyLoadImg = new System.Windows.Forms.PictureBox();
            this.secKeyBtnLoad = new System.Windows.Forms.Button();
            this.secKeyPasswordInput = new System.Windows.Forms.TextBox();
            this.secKeyPasswordLabel = new System.Windows.Forms.Label();
            this.secKeyPathInput = new System.Windows.Forms.TextBox();
            this.secKeyBrowse = new System.Windows.Forms.Button();
            this.secKeyBrowseLabel = new System.Windows.Forms.Label();
            this.secKeyOpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.secKeyLoadImg)).BeginInit();
            this.SuspendLayout();
            // 
            // secKeyLoadErrorLabel
            // 
            this.secKeyLoadErrorLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.secKeyLoadErrorLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyLoadErrorLabel.ForeColor = System.Drawing.Color.Red;
            this.secKeyLoadErrorLabel.Location = new System.Drawing.Point(12, 460);
            this.secKeyLoadErrorLabel.Name = "secKeyLoadErrorLabel";
            this.secKeyLoadErrorLabel.Size = new System.Drawing.Size(405, 65);
            this.secKeyLoadErrorLabel.TabIndex = 0;
            // 
            // secKeyLoadImg
            // 
            this.secKeyLoadImg.ErrorImage = null;
            this.secKeyLoadImg.InitialImage = null;
            this.secKeyLoadImg.Location = new System.Drawing.Point(233, 385);
            this.secKeyLoadImg.Name = "secKeyLoadImg";
            this.secKeyLoadImg.Size = new System.Drawing.Size(54, 54);
            this.secKeyLoadImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.secKeyLoadImg.TabIndex = 6;
            this.secKeyLoadImg.TabStop = false;
            // 
            // secKeyBtnLoad
            // 
            this.secKeyBtnLoad.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(170)))), ((int)(((byte)(240)))));
            this.secKeyBtnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.secKeyBtnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyBtnLoad.ForeColor = System.Drawing.Color.White;
            this.secKeyBtnLoad.Location = new System.Drawing.Point(17, 385);
            this.secKeyBtnLoad.Name = "secKeyBtnLoad";
            this.secKeyBtnLoad.Size = new System.Drawing.Size(188, 54);
            this.secKeyBtnLoad.TabIndex = 4;
            this.secKeyBtnLoad.Text = "Load";
            this.secKeyBtnLoad.UseVisualStyleBackColor = false;
            this.secKeyBtnLoad.Click += new System.EventHandler(this.secKeyBtnLoad_Click);
            // 
            // secKeyPasswordInput
            // 
            this.secKeyPasswordInput.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyPasswordInput.Location = new System.Drawing.Point(16, 289);
            this.secKeyPasswordInput.MinimumSize = new System.Drawing.Size(400, 54);
            this.secKeyPasswordInput.Name = "secKeyPasswordInput";
            this.secKeyPasswordInput.PasswordChar = '●';
            this.secKeyPasswordInput.Size = new System.Drawing.Size(400, 54);
            this.secKeyPasswordInput.TabIndex = 3;
            this.secKeyPasswordInput.UseSystemPasswordChar = true;
            this.secKeyPasswordInput.AutoSize = false;
            // 
            // secKeyPasswordLabel
            // 
            this.secKeyPasswordLabel.AutoSize = true;
            this.secKeyPasswordLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyPasswordLabel.Location = new System.Drawing.Point(12, 248);
            this.secKeyPasswordLabel.Name = "secKeyPasswordLabel";
            this.secKeyPasswordLabel.Size = new System.Drawing.Size(120, 29);
            this.secKeyPasswordLabel.TabIndex = 0;
            this.secKeyPasswordLabel.Text = "Password";
            // 
            // secKeyPathInput
            // 
            this.secKeyPathInput.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyPathInput.Location = new System.Drawing.Point(16, 90);
            this.secKeyPathInput.MinimumSize = new System.Drawing.Size(340, 54);
            this.secKeyPathInput.Name = "secKeyPathInput";
            this.secKeyPathInput.Size = new System.Drawing.Size(400, 54);
            this.secKeyPathInput.TabIndex = 1;
            this.secKeyPathInput.AutoSize = false;
            // 
            // secKeyBrowse
            // 
            this.secKeyBrowse.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(170)))), ((int)(((byte)(240)))));
            this.secKeyBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.secKeyBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyBrowse.ForeColor = System.Drawing.Color.White;
            this.secKeyBrowse.Location = new System.Drawing.Point(17, 150);
            this.secKeyBrowse.Name = "secKeyBrowse";
            this.secKeyBrowse.Size = new System.Drawing.Size(128, 54);
            this.secKeyBrowse.TabIndex = 2;
            this.secKeyBrowse.Text = "Browse";
            this.secKeyBrowse.UseVisualStyleBackColor = false;
            this.secKeyBrowse.Click += new System.EventHandler(this.secKeyBrowse_Click);
            // 
            // secKeyBrowseLabel
            // 
            this.secKeyBrowseLabel.AutoSize = true;
            this.secKeyBrowseLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.secKeyBrowseLabel.Location = new System.Drawing.Point(12, 49);
            this.secKeyBrowseLabel.Name = "secKeyBrowseLabel";
            this.secKeyBrowseLabel.Size = new System.Drawing.Size(321, 29);
            this.secKeyBrowseLabel.TabIndex = 0;
            this.secKeyBrowseLabel.Text = "Load from file (Secretarium file)";
            // 
            // secKeyOpenFileDialog
            // 
            this.secKeyOpenFileDialog.DefaultExt = "secretarium";
            this.secKeyOpenFileDialog.FileName = "YourSecretariumFile";
            this.secKeyOpenFileDialog.Filter = "Secretarium files|*.secretarium|All files|*.*";
            this.secKeyOpenFileDialog.Title = "Load your Secretarium file";
            // 
            // LoadSecKey
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.secKeyLoadErrorLabel);
            this.Controls.Add(this.secKeyLoadImg);
            this.Controls.Add(this.secKeyBtnLoad);
            this.Controls.Add(this.secKeyBrowseLabel);
            this.Controls.Add(this.secKeyPasswordInput);
            this.Controls.Add(this.secKeyBrowse);
            this.Controls.Add(this.secKeyPasswordLabel);
            this.Controls.Add(this.secKeyPathInput);
            this.MinimumSize = new System.Drawing.Size(120, 300);
            this.Name = "LoadSecKey";
            this.Size = new System.Drawing.Size(520, 686);
            ((System.ComponentModel.ISupportInitialize)(this.secKeyLoadImg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label secKeyBrowseLabel;
        private System.Windows.Forms.TextBox secKeyPathInput;
        private System.Windows.Forms.Button secKeyBrowse;
        private System.Windows.Forms.TextBox secKeyPasswordInput;
        private System.Windows.Forms.Label secKeyPasswordLabel;
        private System.Windows.Forms.PictureBox secKeyLoadImg;
        private System.Windows.Forms.Button secKeyBtnLoad;
        private System.Windows.Forms.OpenFileDialog secKeyOpenFileDialog;
        private System.Windows.Forms.Label secKeyLoadErrorLabel;
    }
}
