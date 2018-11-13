using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace WpfApp1
{
    class TextManager
    {

       public static int LettersToNumber(string letters)
       {
            int result = 0;
            if (letters.Length > 0 && letters.All(a => (a >= 'A' && a <= 'Z')))
                try
                {
                    for (int i = letters.Length; i > 0; i--)
                        result += (int)checked(Math.Pow(26, i - 1) + (letters[i - 1] - 'A') *
                            Math.Pow(26, letters.Length - i));
                }
                catch (OverflowException)
                {
                    result = -1;
                }
            else
                result = -1;
            return result;
       }
     
    }
}
