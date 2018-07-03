using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Prospecta.ConnektHub.Services.Translation
{
    public interface ITranslationService
    {
        string GetFieldIdNDescriptionInEnglish(string moduleId);
    }
}
