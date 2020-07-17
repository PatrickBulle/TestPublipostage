using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace TestPublipostage
{
    class Publipostage : IPublipostage
    {
        public void FaitLeTaf(string cheminEtNomDuModele, string cheminEtNomDuDocumentGenere, Dictionary<string, string> dictionnaire, List<int> valeursDuGraphe)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document document = app.Documents.Open(cheminEtNomDuModele);

            foreach (KeyValuePair<string, string> entree in dictionnaire)
            {
                Remplace(document, entree);
            }

            GestionDuGraphe(document, valeursDuGraphe);

            document.ExportAsFixedFormat(cheminEtNomDuDocumentGenere, WdExportFormat.wdExportFormatPDF);
            // Fermer sans enregistrer
            document.Close(false);
            app.Quit();
        }

        private void GestionDuGraphe(Document document, List<int> valeursDuGraphe)
        {
            foreach (InlineShape inlineShape in document.InlineShapes)
            {
                if (inlineShape.HasChart == MsoTriState.msoTrue)
                {
                    Microsoft.Office.Interop.Word.Chart graphique = inlineShape.Chart;

                    Workbook wb = graphique.ChartData.Workbook;
                    Worksheet ws = wb.Worksheets["Feuil1"];

                    ws.Range["B2"].Value = valeursDuGraphe[0];
                    ws.Range["B2"].Value = valeursDuGraphe[1];
                    ws.Range["C3"].Value = valeursDuGraphe[2];
                    ws.Range["D4"].Value = valeursDuGraphe[3];
                    graphique.Refresh();
                    graphique = null;

                    // Pour fermer proprement Excel
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);

                    ws = null;
                    wb = null;
                }
            }

        }

        private void Remplace(Document document, KeyValuePair<string, string> entree)
        {
            foreach (Microsoft.Office.Interop.Word.Range range in document.StoryRanges)
            {
                Find find = range.Find;
                //options
                object matchCase = false;
                object matchWholeWord = true;
                object matchWildCards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = WdReplace.wdReplaceAll;
                object wrap = 1;
                //execute find and replace
                find.Execute(entree.Key, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, entree.Value, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
        }
    }
}
