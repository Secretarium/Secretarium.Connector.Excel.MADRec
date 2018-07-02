namespace Secretarium.Excel
{
    partial class LoadKeys
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
            this.keyPrivateInput = new System.Windows.Forms.TextBox();
            this.keyPublicInput = new System.Windows.Forms.TextBox();
            this.keyLoadErrorLabel = new System.Windows.Forms.Label();
            this.keyLoadImg = new System.Windows.Forms.PictureBox();
            this.keyBtnLoad = new System.Windows.Forms.Button();
            this.keyPrivateLabel = new System.Windows.Forms.Label();
            this.keyPublicLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.keyLoadImg)).BeginInit();
            this.SuspendLayout();
            // 
            // keyPrivateInput
            // 
            this.keyPrivateInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.keyPrivateInput.Location = new System.Drawing.Point(17, 260);
            this.keyPrivateInput.Multiline = true;
            this.keyPrivateInput.Name = "keyPrivateInput";
            this.keyPrivateInput.Size = new System.Drawing.Size(400, 88);
            this.keyPrivateInput.TabIndex = 9;
            // 
            // keyPublicInput
            // 
            this.keyPublicInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.keyPublicInput.Location = new System.Drawing.Point(17, 90);
            this.keyPublicInput.Multiline = true;
            this.keyPublicInput.Name = "keyPublicInput";
            this.keyPublicInput.Size = new System.Drawing.Size(400, 88);
            this.keyPublicInput.TabIndex = 8;
            // 
            // keyLoadErrorLabel
            // 
            this.keyLoadErrorLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.keyLoadErrorLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.keyLoadErrorLabel.ForeColor = System.Drawing.Color.Red;
            this.keyLoadErrorLabel.Location = new System.Drawing.Point(12, 465);
            this.keyLoadErrorLabel.Name = "keyLoadErrorLabel";
            this.keyLoadErrorLabel.Size = new System.Drawing.Size(405, 65);
            this.keyLoadErrorLabel.TabIndex = 7;
            // 
            // keyLoadImg
            // 
            this.keyLoadImg.ErrorImage = null;
            this.keyLoadImg.InitialImage = null;
            this.keyLoadImg.Location = new System.Drawing.Point(236, 398);
            this.keyLoadImg.Name = "keyLoadImg";
            this.keyLoadImg.Size = new System.Drawing.Size(54, 54);
            this.keyLoadImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.keyLoadImg.TabIndex = 6;
            this.keyLoadImg.TabStop = false;
            // 
            // keyBtnLoad
            // 
            this.keyBtnLoad.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(170)))), ((int)(((byte)(240)))));
            this.keyBtnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.keyBtnLoad.ForeColor = System.Drawing.Color.White;
            this.keyBtnLoad.Location = new System.Drawing.Point(16, 398);
            this.keyBtnLoad.Name = "keyBtnLoad";
            this.keyBtnLoad.Size = new System.Drawing.Size(188, 54);
            this.keyBtnLoad.TabIndex = 5;
            this.keyBtnLoad.Text = "Load";
            this.keyBtnLoad.UseVisualStyleBackColor = false;
            this.keyBtnLoad.Click += new System.EventHandler(this.KeyBtnLoad_Click);
            // 
            // keyPrivateLabel
            // 
            this.keyPrivateLabel.AutoSize = true;
            this.keyPrivateLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.keyPrivateLabel.Location = new System.Drawing.Point(12, 219);
            this.keyPrivateLabel.Name = "keyPrivateLabel";
            this.keyPrivateLabel.Size = new System.Drawing.Size(221, 29);
            this.keyPrivateLabel.TabIndex = 3;
            this.keyPrivateLabel.Text = "Base64 Private Key";
            // 
            // keyPublicLabel
            // 
            this.keyPublicLabel.AutoSize = true;
            this.keyPublicLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.keyPublicLabel.Location = new System.Drawing.Point(12, 49);
            this.keyPublicLabel.Name = "keyPublicLabel";
            this.keyPublicLabel.Size = new System.Drawing.Size(214, 29);
            this.keyPublicLabel.TabIndex = 0;
            this.keyPublicLabel.Text = "Base64 Public Key";
            // 
            // LoadKeys
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.keyLoadErrorLabel);
            this.Controls.Add(this.keyPrivateInput);
            this.Controls.Add(this.keyLoadImg);
            this.Controls.Add(this.keyBtnLoad);
            this.Controls.Add(this.keyPublicInput);
            this.Controls.Add(this.keyPublicLabel);
            this.Controls.Add(this.keyPrivateLabel);
            this.MinimumSize = new System.Drawing.Size(120, 300);
            this.Name = "LoadKeys";
            this.Size = new System.Drawing.Size(496, 686);
            ((System.ComponentModel.ISupportInitialize)(this.keyLoadImg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label keyPublicLabel;
        private System.Windows.Forms.Label keyPrivateLabel;
        private System.Windows.Forms.PictureBox keyLoadImg;
        private System.Windows.Forms.Button keyBtnLoad;
        private System.Windows.Forms.Label keyLoadErrorLabel;
        private System.Windows.Forms.TextBox keyPrivateInput;
        private System.Windows.Forms.TextBox keyPublicInput;
    }
}
