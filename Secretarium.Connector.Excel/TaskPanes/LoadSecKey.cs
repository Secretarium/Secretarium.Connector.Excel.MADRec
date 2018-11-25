using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using Secretarium.Helpers;
using System.Drawing;
using ExcelDna.Integration.CustomUI;
using System.ComponentModel;
using System.Security.Cryptography;

namespace Secretarium.Excel
{
    [ComVisible(true)]
    public partial class LoadSecKey : UserControl, ISecretariumCustomTaskPane, INotifyPropertyChanged
    {
        private string _secKeyFilePath = null;
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


        public LoadSecKey()
        {
            InitializeComponent();
            secKeyLoadErrorLabel.DataBindings.Add("Text", this, "ErrorText");
            secKeyLoadImg.DataBindings.Add("Image", this, "LoadImg", true);
        }

        public void InitSecretarium(ExcelRibbon ribbon)
        {
            _xlRibbon = ribbon;
        }

        private void secKeyBrowse_Click(object sender, EventArgs e)
        {
            if (secKeyOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(secKeyOpenFileDialog.FileName))
                {
                    _secKeyFilePath = secKeyOpenFileDialog.FileName;

                    secKeyPathInput.Text = secKeyOpenFileDialog.FileName;
                }
            }
        }

        private void secKeyBtnLoad_Click(object sender, EventArgs e)
        {
            string path = secKeyPathInput.Text;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                ErrorText = "file not found";
                return;
            }

            try
            {
                var cfg = JsonHelper.DeserializeJsonFromFileAs<ScpConfig.KeyConfig>(path);
                cfg.password = secKeyPasswordInput.Text;

                if (cfg.TryGetECDsaKey(out ECDsaCng key))
                {
                    SecretariumFunctions.Scp.Set(key);
                    LoadImg = _xlRibbon.LoadImage("success") as Bitmap;
                    ErrorText = " ";
                }
                else
                {
                    LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                    ErrorText = "Invalid key/password";
                }
            }
            catch (Exception)
            {
                LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                ErrorText = "Unable to load the key, incorrect password ?";
            }
        }
    }
}
