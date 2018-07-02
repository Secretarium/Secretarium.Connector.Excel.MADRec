using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Secretarium.Helpers;
using System.Drawing;
using ExcelDna.Integration.CustomUI;
using System.ComponentModel;

namespace Secretarium.Excel
{
    [ComVisible(true)]
    public partial class LoadKeys : UserControl, ISecretariumCustomTaskPane, INotifyPropertyChanged
    {
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


        public LoadKeys()
        {
            InitializeComponent();
            keyLoadErrorLabel.DataBindings.Add("Text", this, "ErrorText");
            keyLoadImg.DataBindings.Add("Image", this, "LoadImg", true);
        }

        public void InitSecretarium(ExcelRibbon ribbon)
        {
            _xlRibbon = ribbon;
        }

        private void KeyBtnLoad_Click(object sender, EventArgs e)
        {
            string b64PubKey = keyPublicInput.Text;
            string b64PriKey = keyPrivateInput.Text;

            if (string.IsNullOrEmpty(b64PubKey))
            {
                LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                ErrorText = "Missing pulic key";
                return;
            }

            if (string.IsNullOrEmpty(b64PriKey))
            {
                LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                ErrorText = "Missing private key";
                return;
            }

            try
            {
                var key = ECDsaHelper.Import(b64PubKey.FromBase64String(), b64PriKey.FromBase64String());
                SecretariumFunctions.Swss.Set(key);
                LoadImg = _xlRibbon.LoadImage("success") as Bitmap;
                ErrorText = " ";
            }
            catch (Exception)
            {
                try
                {
                    var key = ECDsaHelper.Import(b64PubKey.FromBase64String().ReverseEndianness(), b64PriKey.FromBase64String().ReverseEndianness());
                    SecretariumFunctions.Swss.Set(key);
                    LoadImg = _xlRibbon.LoadImage("success") as Bitmap;
                    ErrorText = " ";
                }
                catch (Exception)
                {
                    LoadImg = _xlRibbon.LoadImage("error") as Bitmap;
                    ErrorText = "Unable to load keys";
                }
            }
        }        
    }
}
