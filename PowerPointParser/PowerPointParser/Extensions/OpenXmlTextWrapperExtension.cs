using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;
using Aaks.PowerPointParser.Html;

namespace Aaks.PowerPointParser.Extensions
{
    public static class OpenXmlTextWrapperExtension
    {
        public static bool IsUnOrderedListItem(this OpenXmlTextWrapper? paragraphWrapper)
        {
            return paragraphWrapper?.PPr?.BuChar?.Char is
                OpenXmlTextModifiers.UlFilledRoundBullet or
                OpenXmlTextModifiers.UlHollowRoundBullet or
                OpenXmlTextModifiers.UlFilledSquareBullet or
                OpenXmlTextModifiers.UlHollowSquareBullet or
                OpenXmlTextModifiers.UlStarBullet or
                OpenXmlTextModifiers.UlArrowBullet or
                OpenXmlTextModifiers.UlCheckmarkBullet;
        }

        public static bool IsOrderedListItem(this OpenXmlTextWrapper? paragraphWrapper)
        {
            return paragraphWrapper?.PPr?.BuAutoNum?.Type is
                OpenXmlTextModifiers.OlArabicPeriod or
                OpenXmlTextModifiers.OlArabicParenRight or
                OpenXmlTextModifiers.OlCapitalRomanNumeralsPeriod or
                OpenXmlTextModifiers.OlCapitalAlphaPeriod or
                OpenXmlTextModifiers.OlLowercaseAlphaRightParen or
                OpenXmlTextModifiers.OlLowerCaseAlphaPeriod or
                OpenXmlTextModifiers.OlLowercaseRomanNumeralsPeriod;
        }
    }
}
