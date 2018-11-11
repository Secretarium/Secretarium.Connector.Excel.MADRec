using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using Secretarium.Helpers;
using System.Security.Cryptography.X509Certificates;
using System.Drawing;
using ExcelDna.Integration.CustomUI;
using System.ComponentModel;
using System.Security.Cryptography;

namespace Secretarium.Excel
{
    [ComVisible(true)]
    public partial class LoadX509 : UserControl, ISecretariumCustomTaskPane, INotifyPropertyChanged
    {
        private string _x509FilePath = null;
        private ExcelRibbon _xlRibbon = null;

        #region Bindings

        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string _errorText;
        public string ErrorText
        {
            get
            {
                return _errorText;
            }
            set
            {
                _errorText = value;
                NotifyPropertyChanged("ErrorText");
            }
        }

        private Image _loadImg;
        public Image LoadImg
        {
            get
            {
                return _loadImg;
            }
            set
            {
                _loadImg = value;
                NotifyPropertyChanged("LoadImg");
            }
        }

        #endregion


        public LoadX509()
        {
            InitializeComponent();
            x509LoadErrorLabel.DataBindings.Add("Text", this, "ErrorText");
            x509LoadImg.DataBindings.Add("Image", this, "LoadImg", true);
        }

        public void InitSecretarium(ExcelRibbon ribbon)
        {
            _xlRibbon = ribbon;
        }

        private void x509Browse_Click(object sender, EventArgs e)
        {
            if (x509OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(x509OpenFileDialog.FileName))
                {
                    _x509FilePath = x509OpenFileDialog.FileName;

                    x509PathInput.Text = x509OpenFileDialog.FileName;
                }
            }
        }

        private void x509BtnLoad_Click(object sender, EventArgs e)
        {
            string path = x509PathInput.Text;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                ErrorText = "Certificate not found";
                return;
            }

            try
            {
                var x509 = X509Helper.LoadX509FromFile(path, x509PasswordInput.Text);
                var key = x509.GetECDsaPrivateKey() as ECDsaCng;
                if (key != null && key.HashAlgorithm == CngAlgorithm.Sha256 && key.KeySize == 256)
                {
                    SecretariumFunctions.Scp.Set(key);
                    LoadImg = _xlRibbon.LoadImage("success") as Bitmap;
                    ErrorText = " ";
                }
                else
                {
                    LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                    ErrorText = "Invalid certificate, expecting ECDSA 256";
                }
            }
            catch (Exception)
            {
                LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                ErrorText = "Unable to load certificate, incorrect password ?";
            }
        }
    }
}
