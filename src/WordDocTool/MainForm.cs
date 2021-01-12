using System;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocTool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        // Starts filling empty cells process.
        private void submitButton_Click(object sender, EventArgs e)
        {
            messageLabel.Text = "";
            progressBar.Visible = true;
            if (!backgroundWorker.IsBusy)
                backgroundWorker.RunWorkerAsync();
            else
                messageLabel.Text = "Busy processing, please wait.";
        }

        // Main bulk of processing.
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string filePath = pathTB.Text;

            // Opens the Word document.
            Word.Application wordApp = new Word.Application();
            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object visible = false;
            try
            {
                Word.Document wordDoc = wordApp.Documents.OpenNoRepairDialog(filePath, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref visible, ref missing, ref missing, ref missing, ref missing);
                wordDoc.Activate();
                foreach (Word.Table table in wordDoc.Tables)
                {
                    int r = table.Rows.Count;
                    for (int i = 1; i <= r; i++)
                    {
                        foreach (Word.Cell cell in table.Rows[i].Cells)
                        {
                            if (cell.Range.Text.Trim() == "\a" || string.IsNullOrWhiteSpace(cell.Range.Text))
                                cell.Range.Text = "N/A" + cell.Range.Text;
                        }


                        // Checks if the cancellation is requested.
                        if (backgroundWorker.CancellationPending)
                        {
                            // Sets Cancel property of DoWorkEventArgs object to true.
                            e.Cancel = true;
                            wordDoc.Save();
                            wordDoc.Close();
                            wordApp.Quit();
                            return;
                        }
                    }
                }
                wordDoc.Save();
                wordDoc.Close();
                wordApp.Quit();
                progressBar.Visible = false;

            }

            // If exception is thrown, must make sure to cancel background process and exit Word application.
            catch (Exception ex) when (ex is NullReferenceException || ex is System.Runtime.InteropServices.COMException)
            {
                progressBar.Visible = false;
                MessageBox.Show("File not found. Exiting program.");
                backgroundWorker.CancelAsync();
                Application.Exit();
            }

            catch (Exception ex)
            {
                progressBar.Visible = false;
                backgroundWorker.CancelAsync();
                progressBar.Visible = false;
                MessageBox.Show("Error. Exiting program.");
                Application.Exit();
            }

            MessageBox.Show("Done!");
        }

        // Outputs appropriate message after process has ended.
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                messageLabel.Text = "Processing cancelled";
            }
            else if (e.Error != null)
            {
                messageLabel.Text = e.Error.Message;
            }
            else
            {
                messageLabel.Text = "Finished.";
            }
            progressBar.Visible = false;
        }

        // Cancels filling document with "N/A" process.
        private void cancelButton_Click(object sender, EventArgs e)
        {
            if (backgroundWorker.IsBusy)
            {
                // Cancel the asynchronous operation if still in progress.
                backgroundWorker.CancelAsync();
                progressBar.Visible = false;
            }
            else
                messageLabel.Text = "No operation in process, cannot cancel.";
        }

        // Opens dialog prompting a user to choose a Word file from the file explorer.
        private void browseButton_Click(object sender, EventArgs e)
        {
            messageLabel.Text = "";
            Thread t = new Thread((ThreadStart)(() => {
            string pathName = "";

                OpenFileDialog openFileDialog = new OpenFileDialog();

                openFileDialog.Filter = "Word Documents (*.doc, *.docx) |*.doc;*.docx";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pathName = openFileDialog.FileName;
                    UpdateTB(pathName);
                }
            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }

        // Fills the file path text box with the path to the Word file chosen by the user.
        private void UpdateTB(string text)
        {
            if (InvokeRequired)
            {
                this.BeginInvoke(new Action<string>(UpdateTB), new object[] { text });
                return;
            }
            pathTB.Text = text;
        }
    }
}
