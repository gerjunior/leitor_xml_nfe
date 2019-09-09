using System;
using System.Windows.Forms;
using XmlNFe.Entities;

namespace XmlNFe
{
    public partial class FormDanfe : Form
    {
        public FormDanfe()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tipExportar.SetToolTip(btnExportar, "Exportar o DataGrid para o Excel");
            tipProcurar.SetToolTip(button2, "Pesquisar o arquivo XML no sistema.");
            tipLer.SetToolTip(btnLerXml, "Fazer a leitura completa do arquivo XML");
        }

        private void btnLerXml_Click(object sender, EventArgs e)
        {
            OpenXML.LerArquivo();
            btnExportar.Enabled = true;

            #region Brute Force

            #region emit
            emit_CNPJ.Text = OpenXML.Dados.Tables["emit"].Rows[0]["CNPJ"].ToString();
            IE.Text = OpenXML.Dados.Tables["emit"].Rows[0]["IE"].ToString();

            UF.Text = OpenXML.Dados.Tables["enderEmit"].Rows[0]["UF"].ToString();
            natOp.Text = OpenXML.Dados.Tables["ide"].Rows[0]["natOp"].ToString();

            #endregion

            #region dest
            xNome.Text = OpenXML.Dados.Tables["dest"].Rows[0]["xNome"].ToString();
            CNPJ.Text = OpenXML.Dados.Tables["dest"].Rows[0]["CNPJ"].ToString();
            xLgr.Text = OpenXML.Dados.Tables["enderDest"].Rows[0]["xLgr"].ToString();
            xBairro.Text = OpenXML.Dados.Tables["enderDest"].Rows[0]["xBairro"].ToString();
            xMun.Text = OpenXML.Dados.Tables["enderDest"].Rows[0]["xMun"].ToString();
            CEP.Text = OpenXML.Dados.Tables["enderDest"].Rows[0]["CEP"].ToString();
            fone.Text = OpenXML.Dados.Tables["enderDest"].Rows[0]["fone"].ToString();
            UF.Text = OpenXML.Dados.Tables["enderDest"].Rows[0]["UF"].ToString();
            dhEmi.Text = Convert.ToDateTime(OpenXML.Dados.Tables["ide"].Rows[0]["dhEmi"].ToString().Substring(0, 10), System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
            dhSaiEnt.Text = Convert.ToDateTime(OpenXML.Dados.Tables["ide"].Rows[0]["dhSaiEnt"].ToString().Substring(0, 10), System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
            dhSaiEntH.Text = Convert.ToDateTime(OpenXML.Dados.Tables["ide"].Rows[0]["dhSaiEnt"].ToString().Substring(11), System.Globalization.CultureInfo.InvariantCulture).ToString("HH:mm:ss");

            #endregion

            #region transp
            xNomeTransp.Text = OpenXML.Dados.Tables["transporta"].Rows[0]["xNome"].ToString();
            xEnderTransp.Text = OpenXML.Dados.Tables["transporta"].Rows[0]["xEnder"].ToString();
            xMunTransp.Text = OpenXML.Dados.Tables["transporta"].Rows[0]["xMun"].ToString();
            UFTransp.Text = OpenXML.Dados.Tables["transporta"].Rows[0]["UF"].ToString();
            pesoB.Text = OpenXML.Dados.Tables["vol"].Rows[0]["pesoB"].ToString();
            pesoL.Text = OpenXML.Dados.Tables["vol"].Rows[0]["pesoL"].ToString();
            marca.Text = OpenXML.Dados.Tables["vol"].Rows[0]["marca"].ToString();
            nVol.Text = OpenXML.Dados.Tables["vol"].Rows[0]["nVol"].ToString();
            modFrete.Text = OpenXML.Dados.Tables["transp"].Rows[0]["modFrete"].ToString();
            #endregion

            #region total
            vBC.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vBC"].ToString();
            vICMS.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vICMS"].ToString();
            vBCST.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vBCST"].ToString();
            vProd.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vProd"].ToString();
            vFrete.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vFrete"].ToString();
            vSeg.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vSeg"].ToString();
            vDesc.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vDesc"].ToString();
            vPIS.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vPIS"].ToString();
            vCONFINS.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vCOFINS"].ToString();
            vOutros.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vOutro"].ToString();
            vNF.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vNF"].ToString();
            vTotTrib.Text = OpenXML.Dados.Tables["ICMSTot"].Rows[0]["vTotTrib"].ToString();
            #endregion
            #region extras
            infCpl.Text = OpenXML.Dados.Tables["infAdic"].Rows[0]["infCpl"].ToString();
            chNFe.Text = OpenXML.Dados.Tables["infProt"].Rows[0]["chNFe"].ToString();
            codBarra.Text = "CÓDIGO DE BARRAS INDISPONÍVEL";
            xMotivo.Text = OpenXML.Dados.Tables["infProt"].Rows[0]["xMotivo"].ToString();
            dhRecbto.Text = Convert.ToDateTime(OpenXML.Dados.Tables["infProt"].Rows[0]["dhRecbto"].ToString().Substring(0, 10), System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
            #endregion

            #endregion

            #region Foreach
            //List<string> tags = new List<string> { "ide", "emit", "dest" }; //??
            //foreach (TextBox _tb in gpDestinatario.Controls.OfType<TextBox>())
            //{
            //    XmlNode node = OpenXML.Documento.DocumentElement;
            //    List<XmlNode> child = node.ChildNodes.>;
            //    foreach (XmlNode node in OpenXML.Documento.DocumentElement)
            //    {
            //        foreach (XmlNode child in node.ChildNodes)
            //        {
            //            foreach (XmlNode grandChild in child.ChildNodes)
            //            {
            //                string nameNode;
            //                if (grandChild.Name == "dest")
            //                if (node.Name == _tb.Name)
            //                {
            //                    _tb.Text = node.InnerText;
            //                }
            //            }
            //        }
            //    }
            //}
            #endregion

            dataGridProdutos.DataSource = OpenXML.Dados.Tables["prod"];

            btnLerXml.Enabled = false;
        }

        private void Button2_Click_1(object sender, EventArgs e)
        {
            OpenXML.AbrirArquivo();
            if (OpenXML.Path != null)
            {
                MessageBox.Show("Destino: " + OpenXML.Path);
                btnLerXml.Enabled = true;
            }
            lblPath.Text = OpenXML.Path;
        }
        private void btnExportar_Click(object sender, EventArgs e)
        {
            OpenXML.ExportarExcel(dataGridProdutos);
        }
    }
}
