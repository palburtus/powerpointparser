using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace Aaks.PowerPointParser.Utils
{
    public class RemoveMalformedHyperlinksRelationshipErrorHandler : RelationshipErrorHandler
    {
        private readonly OpenXmlPackage _package;
        private readonly Dictionary<string, List<string>> _errors;

        public RemoveMalformedHyperlinksRelationshipErrorHandler(OpenXmlPackage package)
        {
            _package = package;
            _errors = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        }

        public override string Rewrite(Uri partUri, string id, string uri)
        {
            var key = partUri.OriginalString
                .Replace("_rels/", string.Empty)
                .Replace(".rels", string.Empty);

            if (!_errors.ContainsKey(key))
            {
                _errors.Add(key, new List<string>());
            }

            _errors[key].Add(id);

            return "http://error";
        }

    }
}


