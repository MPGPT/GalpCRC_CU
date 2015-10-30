using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Net;
using System.Data;

namespace GalpCRC_CU
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //MAIN FUNÇÕES
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Ocurreu uma exeção ao libertar o objecto:" + Environment.NewLine + ex.ToString(), "ERRO");
            }
            finally
            {
                GC.Collect();
            }
        }
        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy-MM-dd.HH;mm;ss");
        }
        public static int GetTimestampH(DateTime value)
        {
            return value.Hour;
        }

        //DIVIDE A LISTA EM 3
        public static List<List<T>> Split<T>(List<T> items, int sliceSize = 3)
        {
            List<List<T>> list = new List<List<T>>();
            for (int i = 0; i < items.Count; i += sliceSize)
                list.Add(items.GetRange(i, Math.Min(sliceSize, items.Count - i)));
            return list;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        //Criação User Gilberto - SVDI
        private void cdugil()
        {
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            string timestamp = GetTimestamp(DateTime.Now);
            string fileName = null;

            if (lbnome.Items.Count == 0)
            { MessageBox.Show("Lista de nomes completos está vazia", "ERRO"); }
            else
            {
                if (cbmercado.SelectedIndex == 0)
                { fileName = path + @"/Templates/CDUML.xlsx"; }
                else if (cbmercado.SelectedIndex == 1)
                { fileName = path + @"/Templates/CDUMR.xlsx"; }
                else
                { MessageBox.Show("Tens de Selecionar um mercado", "ERRO"); return; }

                Excel.Application xlApp = new Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("O excel não está instalado correctamente", "ERRO");
                    return;
                }


                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                var ulist = lbnome.Items.Cast<String>().ToList();
                var idlist = lbid.Items.Cast<String>().ToList();
                var tdlist = lbtipodoc.Items.Cast<String>().ToList();
                var ho = 2;
                var hor = 2;
                var hori = 2;

                foreach (string i in ulist)
                {
                    xlWorkSheet.Cells[ho, 1] = i;
                    ++ho;
                }

                foreach (string i in idlist)
                {
                    xlWorkSheet.Cells[hor, 2] = i;
                    hor++;
                }

                foreach (string i in tdlist)
                {
                    xlWorkSheet.Cells[hori, 3] = i;
                    hori++;
                }

                xlWorkBook.CheckCompatibility = false;
                xlWorkBook.SaveAs(path + @"/criação_de_users_p_svdi_" + timestamp + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, ConflictOption.OverwriteChanges, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                GC.Collect();
            }
        }

        //Criação User GALP ML
        private void cdugalplista0ML()
        {
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            string timestamp = GetTimestamp(DateTime.Now);
            string fileName = null;

            if (lbnome.Items.Count == 0)
            { MessageBox.Show("Lista de nomes completos está vazia", "ERRO"); }
            else
            {
                if (cbmercado.SelectedIndex == 0)
                { fileName = path + @"/Templates/TU_Galpenergia_ML.xlsx"; }
                else if (cbmercado.SelectedIndex == 1)
                { fileName = path + @"/Templates/TU_Galpenergia_MR.xlsx"; }
                else
                { MessageBox.Show("Tens de Selecionar um mercado", "ERRO"); return; }

                Excel.Application xlApp = new Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("O excel não está instalado correctamente", "ERRO");
                    return;
                }


                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                var ulist = lbnome.Items.Cast<String>().ToList();
                var idlist = lbid.Items.Cast<String>().ToList();
                var tdlist = lbtipodoc.Items.Cast<String>().ToList();
                var uolist = lbuseropen.Items.Cast<String>().ToList();
                var ho = 2;
                var hor = 2;
                var hori = 2;
                var horiz = 2;

                foreach (string i in ulist)
                {
                    xlWorkSheet.Cells[ho, 2] = i;
                    ++ho;
                }

                foreach (string i in idlist)
                {
                    xlWorkSheet.Cells[hor, 1] = i;
                    hor++;
                }

                foreach (string i in tdlist)
                {
                    xlWorkSheet.Cells[hori, 3] = i;
                    hori++;
                }

                foreach (string i in uolist)
                {
                    xlWorkSheet.Cells[hori, 4] = i;
                    horiz++;
                }

                xlWorkBook.CheckCompatibility = false;
                xlWorkBook.SaveAs(path + @"/TU_Galpenergia_ML_" + timestamp + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, ConflictOption.OverwriteChanges, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                GC.Collect();
            }
        }

        private void btnlimpartudo_Click(object sender, RoutedEventArgs e)
        {
            tbform.Text = null;
            cbmercado.SelectedItem = null;
            lbid.Items.Clear();
            lbnome.Items.Clear();
            lbtipodoc.Items.Clear();
            lbuseropen.Items.Clear();
        }

        private void lbid_KeyDown(object sender, KeyEventArgs e)
        {
            // Ve se o atalho ctrl.v é clicado e se o clipboard nao ta vazio
            if ((e.Key == Key.V) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))  && Clipboard.ContainsText())
            {
                // Pega no texto do clipboard e separa
                string text = Clipboard.GetText();
                string[] textLines = text.Split(
                    new string[] { Environment.NewLine },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (string i in textLines)
                {
                    lbid.Items.Add(i);
                }

                // Marca o evento como terminado
                e.Handled = true;
            }
        }

        private void lbnome_KeyDown(object sender, KeyEventArgs e)
        {
            // Ve se o atalho ctrl.v é clicado e se o clipboard nao ta vazio
            if ((e.Key == Key.V) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) && Clipboard.ContainsText())
            {
                // Pega no texto do clipboard e separa
                string text = Clipboard.GetText();
                string[] textLines = text.Split(
                    new string[] { Environment.NewLine },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (string i in textLines)
                {
                    lbnome.Items.Add(i);
                }

                // Marca o evento como terminado
                e.Handled = true;
            }
        }

        private void lbtipodoc_KeyDown(object sender, KeyEventArgs e)
        {
            // Ve se o atalho ctrl.v é clicado e se o clipboard nao ta vazio
            if ((e.Key == Key.V) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) && Clipboard.ContainsText())
            {
                // Pega no texto do clipboard e separa
                string text = Clipboard.GetText();
                string[] textLines = text.Split(
                    new string[] { Environment.NewLine },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (string i in textLines)
                {
                    lbtipodoc.Items.Add(i);
                }

                // Marca o evento como terminado
                e.Handled = true;
            }
        }

        private void lbuseropen_KeyDown(object sender, KeyEventArgs e)
        {
            // Ve se o atalho ctrl.v é clicado e se o clipboard nao ta vazio
            if ((e.Key == Key.V) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) && Clipboard.ContainsText())
            {
                // Pega no texto do clipboard e separa
                string text = Clipboard.GetText();
                string[] textLines = text.Split(
                    new string[] { Environment.NewLine },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (string i in textLines)
                {
                    lbuseropen.Items.Add(i);
                }

                // Marca o evento como terminado
                e.Handled = true;
            }
        }

        private void lbid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            lbid.Focus();
        }

        private void lbnome_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            lbnome.Focus();
        }

        private void lbtipodoc_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            lbtipodoc.Focus();
        }

        private void lbuseropen_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            lbuseropen.Focus();
        }

        private void btncu_Click(object sender, RoutedEventArgs e)
        {
            cdugil();
        }
    }
}
