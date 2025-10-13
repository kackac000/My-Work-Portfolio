using static System.Runtime.InteropServices.JavaScript.JSType;

namespace SAP스캐너
{
    public class BarcodeParseResult
    {
        public string? Gtin { get; init; }      // 표준바코드
        public string? Lot { get; init; }       // 제조번호
        public string? Expiry { get; init; }    // YYMMDD (6)
        public string? Serial { get; init; }    // 시리얼넘버
        public string? Value { get; set; }      // 전송값
        public string? Type { get; set; }       // "OBD", "BUNDLE", "GENERAL", "UNKNOWN"
    }
    public static class BarcodeParser
    {
        /// <summary>
        /// 실전 현장 바코드 파싱 (구분자 자동, 대소문자 무시, 모든 분기 포함)
        /// </summary>
        public static BarcodeParseResult Parse(string raw)
        {
            var result = new BarcodeParseResult();

            if (string.IsNullOrWhiteSpace(raw))
            {
                result.Value = "";
                result.Type = "EMPTY";
                return result;
            }

            string sdBarcode = raw.Replace("(", "").Replace(")", "");

            // === checkList/compare3 기준 분기 ===
            string[] checkList = { "8806469024613", "8806469024729", "8806469024910", "추가할숫자2" };
            string compare3 = "8806520009521";
            // 3. 3~15번 인덱스 추출 (총 13글자)
            string subStr = (sdBarcode.Length >= 16) ? sdBarcode.Substring(3, 13) : "";

            // 4. 조건 1) checkList에 포함되어 있으면 15번 인덱스 뒤(=16번째 위치)에 "21" 추가
            if (checkList.Contains(subStr))
                sdBarcode = sdBarcode.Insert(16, "21");
            // 5. 조건 2) subStr이 compare3와 같으면 "11241010" 모두 삭제
            if (subStr == compare3)
                sdBarcode = sdBarcode.Replace("11241010", "");

            // === 출고 OBD ===
            if ((sdBarcode.StartsWith("0") && sdBarcode.Length == 10) ||
                (sdBarcode.StartsWith("8") && sdBarcode.Length == 9) ||
                (sdBarcode.StartsWith("1") && sdBarcode.Length == 9) ||
                (sdBarcode.StartsWith("5") && sdBarcode.Length == 8))
            {
                result.Value = sdBarcode;
                result.Type = "OBD";
                return result;
            }

            string[] delimiters = { "]d2", "]D2", "<GS>", "<gs>", "[gs]", "[GS]" };     // 묶음 바코드 분기
            string sdBarcodeLower = sdBarcode.ToLower();                // 대문자를 소문자로 변경
            string? foundDelim = delimiters.FirstOrDefault(d => sdBarcodeLower.Contains(d.ToLower())) ?? "";
            int len = sdBarcode.Length;                                 // 바코드 전체 문자열 길이
            int idx = sdBarcodeLower.IndexOf(foundDelim.ToLower());     // 바코드의 구분자 까지의 개수

            //MessageBox.Show(sdBarcodeLower.ToString());
            //MessageBox.Show(foundDelim.ToString());
            //MessageBox.Show(idx.ToString());
            //MessageBox.Show(sdBarcode.Substring(idx + foundDelim.Length, 2).ToString());

            // 1. 'GENERAL_SPECIAL' 타입 조건
            if (idx == 0 && sdBarcode.Length <= 17 && (
                    sdBarcode.StartsWith("01088") ||
                    sdBarcode.StartsWith("188") ||
                    sdBarcode.StartsWith("088") ||
                    sdBarcode.StartsWith("88")
                ))
            {
                result.Type = "GENERAL_SPECIAL";
                result.Value = sdBarcode;
                return result;
            }
            // === 개별 - 단순 (제조/유효/시리얼 없음)
            else if (
            (sdBarcode.StartsWith("01088") ||
            sdBarcode.StartsWith("88") ||
            sdBarcode.StartsWith("3") ||
            ((sdBarcode.Length >= 3 && sdBarcode[0] >= 0 && sdBarcode[0] >= 9) && sdBarcode[1] == '8' && sdBarcode[2] == '8')) &&
            idx == 0    //(sdBarcode.Length >= 13 && sdBarcode.Length <= 25)
            )                                           
            {
                // 개별 - 단순 처리
                result.Type = "GENERAL";
                result.Value = sdBarcode;
                
                //MessageBox.Show(result.Type.ToString());
                
                return result;
            }
            // === 개별 - 복잡 (제조/유효/시리얼 있음)
            else if (
            sdBarcode.StartsWith("010") && idx >= 18     // sdBarcode.Length <= 26
            )                                                                                   
            {

                //if (sdBarcode.Substring(16, 2) == "21" && idx >= 19)
                //{
                //    sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(16, 2)}){sdBarcode.Substring(18, idx - 18)}";
                //    // result.Serial 시리얼넘버 입력 분

                //    if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "17")
                //    {
                //        // result.Expiry 유효기한 입력 부분
                //        if (sdBarcode.Substring(idx + foundDelim.Length + 6, 2) == "10")
                //        {
                //            // result.Lot 제조번호 입력부분
                //        }
                //    }
                //    else if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "10")
                //    {
                //        // result.Lot 제조번호 입력부분
                //    }
                //}
                //else if (sdBarcode.Substring(16, 2) == "10")
                //{
                //    // result.Lot 제조번호 입력부분
                //    if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "17")
                //    {
                //        // result.Expiry 유효기한 입력 부분
                //        if (sdBarcode.Substring(idx + foundDelim.Length + 6, 2) == "21")
                //        {
                //            // result.Serial 시리얼넘버 입력 분
                //            sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(idx + foundDelim.Length, 2)}){sdBarcode.Substring(idx + foundDelim.Length + 6)}";
                //        }
                //    }
                //    else if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "21")
                //    {
                //        // result.Lot 제조번호 입력부분
                //    }
                //}
                //else if (sdBarcode.Substring(16, 2) == "17")
                //{
                //    // result.Expiry 제조번호 입력부분
                //    if (sdBarcode.Substring(24, 2) == "10")
                //    {
                //        if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "21")
                //        {
                //            // result.Serial 시리얼넘버 입력 분
                //            sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(idx + foundDelim.Length, 2)}){sdBarcode.Substring(idx + foundDelim.Length + 2)}";
                //        }
                //        else if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "10")
                //        {
                //            // result.Lot 제조번호 입력부분
                //        }
                //    }
                //    else if (sdBarcode.Substring(24, 2) == "21")
                //    {

                //    }
                //}
                // 개별 - 복잡 처리
                result.Type = "GENERAL";
                result.Value = sdBarcode;

                //MessageBox.Show(result.Type.ToString());

                return result;
            }
            // === 묶음 - 단순 (제조/유효/시리얼 없음)
            else if (
            (sdBarcode.StartsWith("000") ||
            (sdBarcode.Length > 2 && sdBarcode[0] == '0' && sdBarcode[1] == '0' && (sdBarcode[2] >= '1' && sdBarcode[2] <= '9')) ||
            (sdBarcode.Length > 2 && sdBarcode[0] == '0' && sdBarcode[1] == '1' && (sdBarcode[2] >= '1' && sdBarcode[2] <= '9'))) &&
            idx == 0        // sdBarcode.Length <= 20
            )                                          
            {
                // 묶음 - 단순 처리
                result.Type = "BUNDLE";
                result.Value = sdBarcode;

                //MessageBox.Show(result.Type.ToString());

                return result;
            }
            // === 묶음 - 복잡 (제조/유효/시리얼 있음)
            else if (
            (sdBarcode.StartsWith("000") ||
            (sdBarcode.Length > 2 && sdBarcode[0] == '0' && sdBarcode[1] == '0' && (sdBarcode[2] >= '1' && sdBarcode[2] <= '9')) ||
            (sdBarcode.Length > 2 && sdBarcode[0] == '0' && sdBarcode[1] == '1' && (sdBarcode[2] >= '1' && sdBarcode[2] <= '9'))) &&
            idx >= 18        // sdBarcode.Length >= 21
            )                                           
            {
                // 묶음 - 복잡 처리
                result.Type = "BUNDLE";

                if(sdBarcode.Substring(16, 2) == "21")
                {
                    sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(16, 2)}){sdBarcode.Substring(18, idx - 18)}";
                }
                else if (sdBarcode.Substring(16, 2) == "17" || sdBarcode.Substring(16, 2) == "10")
                {
                    if (sdBarcode.Substring(24, 2) == "21")
                    {
                        sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(22, 2)}){sdBarcode.Substring(24,idx-18)}";
                    }
                    else if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "21")
                    {
                        sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(idx + foundDelim.Length, 2)}){sdBarcode.Substring(idx + foundDelim.Length + 2)}";
                    }
                    else if (sdBarcode.Substring(idx + foundDelim.Length, 2) == "17") 
                    {
                        sdBarcode = $"({sdBarcode.Substring(0, 2)}){sdBarcode.Substring(2, 14)}({sdBarcode.Substring(idx + foundDelim.Length + 8, 2)}){sdBarcode.Substring(idx + foundDelim.Length + 10)}";
                    }
                }

                    result.Value = sdBarcode;
                return result;
            }
            else
            {
                // 나머지 처리 (UNKNOWN 등)
                // 기타 예외/미처리 케이스도 원본 반환
                result.Value = sdBarcode;
                result.Type = "UNKNOWN";
                return result;
            }
        }
    }
}


