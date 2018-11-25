using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Secretarium.Helpers;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Forms;
using Xl = Microsoft.Office.Interop.Excel;
using XlCore = Microsoft.Office.Core;

namespace Secretarium.Excel
{
    [ComVisible(true)]
    public class SecretariumRibbon : ExcelRibbon
    {
        private CustomTaskPane _secKeyTaskPane;
        private CustomTaskPane _x509TaskPane;
        private CustomTaskPane _keysTaskPane;
        
        public override string GetCustomUI(string RibbonID)
        {
            return @"
                <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' xmlns:Q='Secretarium' loadImage='LoadImage'>
                  <ribbon>
                    <tabs>
                      <tab idQ='Q:Secretarium' label='Secretarium'>
                        <group idQ='Q:SecretariumGrpConnect' label='Secretarium'>
                          <splitButton id='loadKeySplitBtn' size='large'>
                            <button 
                              id='loadKeyBtn' image='secretarium' label='Load keys'
                              onAction='ShowSecKeyTaskPane' />
                            <menu id='loadKeySplitBtnMenu' itemSize='large'>
                              <button 
                                id='loadKeyBtn1' label='Load secretarium key'
                                image='secretarium'
                                description='Load identification and signature keys from a Secretarium file'
                                onAction='ShowSecKeyTaskPane' />
                              <button 
                                id='loadKeyBtn2' label='Load from encoded keys'
                                imageMso='FileDocumentEncrypt'
                                description='Load identification and signature keys from base 64 encoded strings'
                                onAction='ShowBase64KeysTaskPane' />
                              <button 
                                id='loadX509Btn' label='Load from X509'
                                imageMso='SignatureShow'
                                description='Load identification and signature keys from X509 certificate'
                                onAction='ShowX509TaskPane' />
                            </menu>
                          </splitButton>
                          <button 
                            id='connect' label='Connect' 
                            imageMso='ServerConnection' size='large'
                            screentip ='Connect to Secretarium'
                            supertip='Opens a secure connection to Secretarium'
                            onAction='Connect' />
                        </group>
                        <group idQ='Q:SecretariumUtils' label='Utils'>
                          <splitButton id='forceRecomputeSplitBtn' size='large'>
                            <button 
                              id='forceRecomputeBtn' imageMso='CalculateNow' label='Calculate selection'
                              onAction='CalculateSelection' />
                            <menu id='forceRecomputeSplitBtnMenu' itemSize='large'>
                              <button 
                                id='forceRecomputeBtn1' label='Calculate selection'
                                imageMso='CalculateNow'
                                description='Recompute selected cells'
                                onAction='CalculateSelection' />
                              <button 
                                id='forceRecomputeBtn2' label='Force calculate selection'
                                imageMso='CalculateFull'
                                description='Force recomputation of selected cells'
                                onAction='ForceCalculateSelection' />
                            </menu>
                          </splitButton>
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void ShowSecKeyTaskPane(IRibbonControl control)
        {
            if (_secKeyTaskPane == null)
            {
                var secKeyLoader = new LoadSecKey();
                secKeyLoader.InitSecretarium(this);

                _secKeyTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(secKeyLoader, "Secretarium");
                _secKeyTaskPane.Visible = true;
                _secKeyTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _secKeyTaskPane.Width = 500;
            }
            else
            {
                _secKeyTaskPane.Visible = true;
            }
        }
        public void ShowX509TaskPane(IRibbonControl control)
        {
            if (_x509TaskPane == null)
            {
                var x509Loader = new LoadX509();
                x509Loader.InitSecretarium(this);

                _x509TaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(x509Loader, "Secretarium");
                _x509TaskPane.Visible = true;
                _x509TaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _x509TaskPane.Width = 500;
            }
            else
            {
                _x509TaskPane.Visible = true;
            }
        }
        public void ShowBase64KeysTaskPane(IRibbonControl control)
        {
            if (_keysTaskPane == null)
            {
                var keysLoader = new LoadKeys();
                keysLoader.InitSecretarium(this);

                _keysTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(keysLoader, "Secretarium");
                _keysTaskPane.Visible = true;
                _keysTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _keysTaskPane.Width = 500;
            }
            else
            {
                _keysTaskPane.Visible = true;
            }
        }

        public void Connect(IRibbonControl control)
        {
            ECDsaCng key = null;

            if (string.IsNullOrEmpty(SecretariumFunctions.Scp.PublicKey))
            {
                MessageBox.Show("Please register you identity first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if(key != null && key.PublicKey().ToBase64String() != SecretariumFunctions.Scp.PublicKey)
                SecretariumFunctions.Scp.Set(key);

            if (SecretariumFunctions.Scp.State.IsClosed())
                SecretariumFunctions.Scp.Connect();
        }

        public void CalculateSelection(IRibbonControl control)
        {
            var app = (Xl.Application)ExcelDnaUtil.Application;
            var selection = app.Selection as Xl.Range;
            if (selection == null)
                return;

            selection.Calculate();
        }
        public void ForceCalculateSelection(IRibbonControl control)
        {
            var app = (Xl.Application)ExcelDnaUtil.Application;
            var selection = app.Selection as Xl.Range;
            if (selection == null)
                return;

            foreach(Xl.Range c in selection.Cells)
            {
                var f = c.FormulaR1C1;
                c.FormulaR1C1 = "'" + f;
                c.FormulaR1C1 = f;
            }
        }
    }

    public interface ISecretariumCustomTaskPane
    {
        void InitSecretarium(ExcelRibbon ribbon);
    }
}
