using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;

namespace exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Criacao do documento
                // Cria um documento com o nome exemploDoc
                Document exemploDoc = new Document();
            #endregion

            #region Criacao de secao no documento
                // Adiciona uma seção com o nome secaoCapa ao documento
                // Cada secao pode ser entendida como uma pagina do documento
                Section secaoCapa = exemploDoc.AddSection();
            #endregion

            #region Criar um paragrafo
                // Cria um paragrafo com o nome titulo e adiciona à seção secaoCapa
                // Os paragrafos são necessários para inserção de textos, imagens, 
                // tabelas etc
                Paragraph titulo = secaoCapa.AddParagraph();
            #endregion

            #region Adiciona texto ao paragrafo
                // Adiciona o texto Exemplo de titulo ao paragrafo titulo
                titulo.AppendText("Exemplo de título\n\n");
            #endregion

            #region Formatar paragrafo
                // Através da propriedade HorizontalAlignment
                // é possível alinhar o parágrafo
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

                // Cria um estilo com o nome estilo01 e adiciona ao documento
                ParagraphStyle estilo01 = new ParagraphStyle(exemploDoc);

                // Adiciona um nome ao estilo01
                estilo01.Name = "Cor do titulo";

                // Definir a cor do texto
                estilo01.CharacterFormat.TextColor = Color.DarkBlue;

                // Define que o texto será em negrito
                estilo01.CharacterFormat.Bold = true;

                // Adiciona o estilo01 ao documento exemploDoc
                exemploDoc.Styles.Add(estilo01);

                // Aplica o estilo01 ao parágrafo titulo
                titulo.ApplyStyle(estilo01.Name);
            #endregion

            #region Trabalhar com tabulação
                // Adiciona um paragrafo textoCapa à seção secaoCapa
                Paragraph textoCapa = secaoCapa.AddParagraph();

                // Adiciona um texto ao parágrfo com tabulação
                textoCapa.AppendText("\tEste é um exemplo de texto com tabulação\n");

                // Adiciona um novo parágrafo à mesma seção (secaoCapa)
                Paragraph textoCapa2 = secaoCapa.AddParagraph();

                // Adiciona um texto ao parágrafo textoCapa2 com concatenação
                textoCapa2.AppendText("\tBasicamente, então, uma seção representa uma página do documento e os parágrafos dentro de uma mesma seção," + "obviamente, aparecem na mesma página");
            #endregion
        }
    }
}
