using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAP스캐너
{
    internal class ScriptBarcodeHandler
    {
        public bool TryProcessBarcode(IntPtr targetHwnd, string barcode)
        {
            // TODO: 나중에 스크립트 자동화 실제 코드 구현
            return false; // 지금은 항상 실패(즉시 키마 분기로 빠짐)
        }
    }
}
