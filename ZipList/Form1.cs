using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZipList
{
    public partial class Form1 : Form
    {
        private string zipFile = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openDialog = new OpenFileDialog();
            openDialog.Filter = "ZIP File|*.zip";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                zipFile = openDialog.FileName;
                zipList(zipFile);
                Text = Path.GetFileName(zipFile);
            }
        }

        private void zipList(string fileName)
        {
            listView1.Items.Clear();
            using (var zipArc = ZipFile.OpenRead(fileName))
            {
                zipArc.Entries.ToList().ForEach(entry =>
                {
                    var item = new ListViewItem(new string[]
                    {
                        entry.FullName,
                        entry.Length.ToString("#,0"),
                        entry.CompressedLength.ToString("#,0"),
                        entry.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss")
                    });
                    listView1.Items.Add(item);
                });
            }
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (listView1.SelectedItems.Count > 0)
                {
                    contextMenuStrip2.Show(listView1, e.Location);
                }
                else
                {
                    contextMenuStrip1.Show(listView1, e.Location);
                }
            }
        }

        private void extractToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count <= 0)
            {
                return;
            }
            var saveDialog = new SaveFileDialog();
            if (saveDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            var list = new List<ListViewItem>();
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                list.Add(item);
            }
            var saveFolder = Path.GetDirectoryName(saveDialog.FileName);
            using (var zipArc = ZipFile.OpenRead(zipFile))
            {
                list.ForEach(item =>
                {
                    var entry = zipArc.GetEntry(item.SubItems[0].Text);
                    entry.ExtractToFile(Path.Combine(saveFolder, entry.Name), true);
                });
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (zipFile == string.Empty)
            {
                newZipFileCreateAdd();
            }
            zipFileAdd();
        }

        private async void zipFileAdd()
        {
            var openDialog = new OpenFileDialog();
            openDialog.Multiselect = true;
            openDialog.Filter = "All File|*.*";
            openDialog.Title = "追加するファイル指定";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                await Task.Run(() =>
                {
                    // オンメモリで操作する関係上、ファイル単位でOpen～書き込みを行う
                    openDialog.FileNames.ToList().ForEach(file =>
                    {
                        //読み取りと書き込みができるようにして、ZIP書庫を開く
                        using (ZipArchive a = ZipFile.Open(zipFile, ZipArchiveMode.Update, Encoding.GetEncoding("shift_jis")))
                        {
                            //ファイル「C:\test\1.txt」を「1.txt」としてZIPに追加する
                            var e = a.CreateEntryFromFile(file, Path.GetFileName(file));
                            var fl = new FileInfo(file);
                            var item = new ListViewItem(new string[]
                            {
                                e.FullName,
                                fl.Length.ToString("#,0"),
                                0.ToString("#,0"),
                                e.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss")
                            });
                            
                            this.Invoke((MethodInvoker)(() => listView1.Items.Add(item)));
                        }
                    });
                });
            }
        }

        private void newZipFileCreateAdd()
        {
            var saveZipDialog = new SaveFileDialog();
            saveZipDialog.Filter = "ZIP File|*.zip";
            if (saveZipDialog.ShowDialog() == DialogResult.OK)
            {
                zipFile = saveZipDialog.FileName;
                Text = Path.GetFileName(zipFile);
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var list = new List<ListViewItem>();
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                list.Add(item);
            }
            using (var zipArc = ZipFile.Open(zipFile, ZipArchiveMode.Update))
            {
                list.ForEach(item =>
                {
                    var entry = zipArc.GetEntry(item.Text);
                    entry.Delete();
                    listView1.Items.Remove(item);
                });
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            var item = listView1.SelectedItems[0];
            var file = Path.Combine(System.IO.Path.GetTempPath(), item.Text);
            using (var zipArc = ZipFile.Open(zipFile, ZipArchiveMode.Read))
            {
                var entry = zipArc.GetEntry(item.Text);
                entry.ExtractToFile(file, true);
                var p = Process.Start(file);
            }
        }
    }
}
