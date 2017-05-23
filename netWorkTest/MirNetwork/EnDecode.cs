using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace netWorkTest.MirNetwork
{
    class EnDecode
    {
        static  byte []Decode6BitMask = { 0xfc, 0xf8, 0xf0, 0xe0, 0xc0 };
        public static int  fnEncode6BitBufA(byte[] pszSrc, byte[] pszDest)
        {
            int nDestPos = 0;
            int nRestCount = 0;
            int chMade = 0, chRest = 0;

            for (int i = 0; i < pszSrc.Length; i++)
            {
                if (nDestPos >= pszDest.Length) break;

                chMade = ((chRest | (pszSrc[i] >> (2 + nRestCount))) & 0x3f);
                chRest = (((pszSrc[i] << (8 - (2 + nRestCount))) >> 2) & 0x3f);

                nRestCount += 2;

                if (nRestCount < 6)
                    pszDest[nDestPos++] = (byte)(chMade + 0x3c);
                else
                {
                    if (nDestPos < pszDest.Length - 1)
                    {
                        pszDest[nDestPos++] = (byte)(chMade + 0x3c);
                        pszDest[nDestPos++] = (byte)(chRest + 0x3c);
                    }
                    else
                        pszDest[nDestPos++] = (byte)(chMade + 0x3c);

                    nRestCount = 0;
                    chRest = 0;
                }
            }

            if (nRestCount > 0)
                pszDest[nDestPos++] = (byte)(chRest + 0x3c);

            //	pszDest[nDestPos] = '\0';

            return nDestPos;
        }
        public static int  fnDecode6BitBufA(byte[] pszSrc, byte[] pszDest,int offsetDest, int nDestLen)
        {
            int nLen = pszSrc.Length;//memlen((const char *)pszSrc) - 1;
            int nDestPos = offsetDest, nBitPos = 2;
            int nMadeBit = 0;
            int ch, chCode, tmp=0;

            for (int i = 0; i < nLen; i++)
            {
                if ((pszSrc[i] - 0x3c) >= 0)
                    ch = pszSrc[i] - 0x3c;
                else
                {
                    nDestPos = 0;
                    break;
                }

                if (nDestPos >= nDestLen) break;

                if ((nMadeBit + 6) >= 8)
                {
                    chCode = (tmp | ((ch & 0x3f) >> (6 - nBitPos)));
                    //New model begin LOWORD HIBYTE LOWORD LOBYTE
                    //chCode=chCode^(HIBYTE(LOWORD(0x0C08BA52E))+LOBYTE(LOWORD(0x0C08BA52E)));
                    //chCode=chCode^LOBYTE(LOWORD(0x408D4D));
                    //chCode=Decode6BitMask[nBitPos - 2]^LOBYTE(0x8D34);
                    //New model end


                    pszDest[nDestPos++] = (byte)chCode;

                    nMadeBit = 0;

                    if (nBitPos < 6)
                        nBitPos += 2;
                    else
                    {
                        nBitPos = 2;
                        continue;
                    }
                }

                tmp = ((ch << nBitPos) & Decode6BitMask[nBitPos - 2]);

                nMadeBit += (8 - nBitPos);
            }

            //pszDest[nDestPos] = '\0';

            return nDestPos;
        }
    }
}
