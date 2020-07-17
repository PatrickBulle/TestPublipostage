using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TestPublipostage
{
    class Program
    {
        static void Main()
        {
            const string modele = @"C:\Users\PBulle\Desktop\ModèleLettreGagnant.docx";
            const string documentGenere = @"C:\Users\PBulle\Desktop\Resultat.pdf";
            Dictionary<string, string> dictionnaire = new Dictionary<string, string>();
            List<int> valeurs = new List<int>();

            dictionnaire.Add("#VILLE_ENTETE#", "Besançon");
            dictionnaire.Add("#DATE_ENTETE#", "Jeudi 16 Juillet 2020");
            dictionnaire.Add("#CIVILITE#", "Monsieur");
            dictionnaire.Add("#NOM#", "GARNIER");
            dictionnaire.Add("#PRENOM#", "Yan");

            for (int i = 0; i < 12; i++)
            {
                valeurs.Add(i);
            }
            
            IPublipostage pub = new Publipostage();
            pub.FaitLeTaf(modele, documentGenere, dictionnaire, valeurs);
        }
    }
}
