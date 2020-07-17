using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestPublipostage
{
    public interface IPublipostage
    {
        void FaitLeTaf(string cheminEtNomDuModele, string cheminEtNomDuDocumentGenere, Dictionary<string, string> correspodances, List<int> valeursDuGraphe);
    }
}
