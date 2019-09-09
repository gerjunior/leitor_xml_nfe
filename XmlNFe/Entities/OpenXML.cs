using System;
using System.Data;
using System.Windows.Forms;

namespace XmlNFe.Entities
{
    static class OpenXML
    {
        public static DataSet Dados = new DataSet();
        public static string Path { get; set; }

        public static void AbrirArquivo()
        {
            OpenFileDialog theDialog = new OpenFileDialog
            {
                Title = "Abrir arquivo XML",
                Filter = "XML files|*.xml",
                InitialDirectory = @"C:\"
            };

            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                Path = theDialog.FileName.ToString();
            }
        }
        public static void LerArquivo()
        {
            Dados.ReadXml(Path);
        }
        public static void ExportarExcel(DataGridView dataGrid)
        {
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();

            if (dataGrid.Rows.Count > 0)
            {
                try
                {
                    XcelApp.Application.Workbooks.Add(Type.Missing);
                    for (int i = 1; i < dataGrid.Columns.Count + 1; i++)
                    {
                        XcelApp.Cells[1, i] = dataGrid.Columns[i - 1].HeaderText;
                    }
                    for (int i = 0; i < dataGrid.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGrid.Columns.Count; j++)
                        {
                            XcelApp.Cells[i + 2, j + 1] = dataGrid.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    XcelApp.Columns.AutoFit();
                    XcelApp.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro : " + ex.Message);
                    XcelApp.Quit();
                }
            }
        }
    }
}
