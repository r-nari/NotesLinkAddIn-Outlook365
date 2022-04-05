using System;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;

namespace NotesLinkAddIn_x64
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    を Outlook のシャットダウン時に実行する必要があります。https://go.microsoft.com/fwlink/?LinkId=506785 をご覧ください
        }
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new Ribbon1() }
            );
        }

        public static ThisAddIn instance = null;
        public static ThisAddIn Instance()
        {
            return instance;
        }

        private void setClipboardText(object args)
        {
            Clipboard.SetText((String)args);
        }

        private void isEdited()
        {
            MessageBox.Show("Edited");
        }

        String raw_link = "";
        String notes_link = "";
        String pattern = "[-<:> ]";
        String[] arr = { };
        int index = -1;
        bool flag = false;
        bool onProcessing = false;
        IDataObject iData = null;

        internal async void onButtonNotesLink()
        {
            if (onProcessing)
            {
                return;
            }
            onProcessing = true;
            iData = Clipboard.GetDataObject();
            if (iData.GetDataPresent(DataFormats.Text, false))
            {
                raw_link = (String)iData.GetData(DataFormats.Text);
            }
            else
            {
                MessageBox.Show("No text data was found in the clipboard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                onProcessing = false;
                return;
            }
            arr = Regex.Split(raw_link, pattern);
            index = -1;
            flag = false;
            foreach (String value in arr)
            {
                index++;
                if (value == "NDL")
                {
                    flag = true;
                    break;
                }
            }
            notes_link = "Notes://";
            if (arr.Length > index + 19 && flag && Regex.Split(arr[index + 19], "[/=]").Length > 1)
            {
                notes_link = notes_link + Regex.Split(arr[index + 19], "[/=]")[1] + "/" + arr[index + 3] + arr[index + 4] + "/0/" + arr[index + 13].Substring(2) + arr[index + 14] + arr[index + 15].Substring(2) + arr[index + 16];
                ClipboardHelper.CopyToClipboard("<a href=\"" + notes_link + "\">NotesLink</a>", notes_link);
                SendKeys.Send("^v");
                await Task.Delay(1000);
                Thread t = new Thread(new ParameterizedThreadStart(setClipboardText));
                t.SetApartmentState(ApartmentState.STA);
                t.Start(raw_link);
                t.Join();
            }
            else
            {
                MessageBox.Show("No NotesLink data was found in the clipboard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            onProcessing = false;
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
