# 🚀 Google Sheets Amount to Words (No Script)

A powerful, native Google Sheets formula to convert numerical amounts into words (text). 
It supports **5 languages** and **16 currencies**, handles grammar (gender, plurals), and processes decimal parts (cents/kopecks/pence).

> **Key Feature:** It uses `LET` and `LAMBDA` functions. **No Apps Script required.** It works instantly and doesn't trigger security warnings.

![Google Sheets Badge](https://img.shields.io/badge/Google%20Sheets-Native%20Formula-green)
![License](https://img.shields.io/badge/license-MIT-blue)

## ✨ Features

* **⚡ Zero Scripts:** Pure formula logic. Blazing fast recalculations.
* **🌍 5 Languages:** English (US/UK), German (DE), Spanish (ES), Ukrainian (UA), Russian (RU).
* **💰 16 Currencies:** USD, EUR, UAH, GBP, JPY, CHF, and 10 more.
* **🧠 Grammar Aware:** * Handles currency gender (e.g., "One Dollar" vs "Одна Гривня").
    * Handles pluralization rules for Slavic languages (1, 2-4, 5+).
    * Handles German compound numbers (e.g., "einundzwanzig").
* **🪙 Smart Decimals:** Converts cents, pence, kopecks, etc., respecting their gender and grammar.

## 📦 Installation & Usage

There are two ways to use this. The recommended way is creating a **Named Function**.

### Method 1: Named Function (Recommended)
This keeps your cells clean: `=AMOUNT_TO_WORDS(A2, "USD", "EN")`.

1.  Open your Google Sheet.
2.  Go to **Data** > **Named functions**.
3.  Click **Add new function**.
4.  **Function name:** `AMOUNT_TO_WORDS`
5.  **Argument placeholders:**
    * `val`
    * `curr_code`
    * `target_lang`
6.  **Formula definition:** Copy and paste the code below into the definition box.

<details>
<summary><b>Click to view the Source Code</b></summary>

```excel
=LET(
    val, val,
    curr_code, UPPER(curr_code),
    target_lang, UPPER(target_lang),

    comment_numerals, "--- 1. NUMERALS DATA (DO NOT EDIT) ---",
    RAW_NUM_STR, SWITCH(target_lang,
        "UA", "|один|два|три|чотири|п'ять|шість|сім|вісім|дев'ять~|одна|дві|три|чотири|п'ять|шість|сім|вісім|дев'ять~десять|одинадцять|дванадцять|тринадцять|чотирнадцять|п'ятнадцять|шістнадцять|сімнадцять|вісімнадцять|дев'ятнадцять~||двадцять|тридцять|сорок|п'ятдесят|шістдесят|сімдесят|вісімдесят|дев'яносто~|сто|двісті|триста|чотириста|п'ятсот|шістьсот|сімсот|вісімсот|дев'ятсот~тисяча|тисячі|тисяч|1~мільйон|мільйони|мільйонів|0~мільярд|мільярди|мільярдів|0~Нуль",
        "RU", "|один|два|три|четыре|пять|шесть|семь|восемь|девять~|одна|две|три|четыре|пять|шесть|семь|восемь|девять~десять|одиннадцать|двенадцать|тринадцать|четырнадцать|пятнадцать|шестнадцать|семнадцать|восемнадцать|девятнадцать~||двадцать|тридцать|сорок|пятьдесят|шестьдесят|семьдесят|восемьдесят|девяносто~|сто|двести|триста|четыреста|пятьсот|шестьсот|семьсот|восемьсот|девятьсот~тысяча|тысячи|тысяч|1~миллион|миллиона|миллионов|0~миллиард|миллиарда|миллиардов|0~Ноль",
        "DE", "|eins|zwei|drei|vier|fünf|sechs|sieben|acht|neun~|eine|zwei|drei|vier|fünf|sechs|sieben|acht|neun~zehn|elf|zwölf|dreizehn|vierzehn|fünfzehn|sechzehn|siebzehn|achtzehn|neunzehn~||zwanzig|dreißig|vierzig|fünfzig|sechzig|siebzig|achtzig|neunzig~|einhundert|zweihundert|dreihundert|vierhundert|fünfhundert|sechshundert|siebenhundert|achthundert|neunhundert~tausend|tausend|tausend|0~Million|Millionen|Millionen|1~Milliarde|Milliarden|Milliarden|1~Null",
        "ES", "|un|dos|tres|cuatro|cinco|seis|siete|ocho|nueve~|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve~diez|once|doce|trece|catorce|quince|dieciséis|diecisiete|dieciocho|diecinueve~||veinte|treinta|cuarenta|cincuenta|sesenta|setenta|ochenta|noventa~|ciento|doscientos|trescientos|cuatrocientos|quinientos|seiscientos|setecientos|ochocientos|novecientos~mil|mil|mil|0~millón|millones|millones|0~mil millones|mil millones|mil millones|0~Cero",
        "|One|Two|Three|Four|Five|Six|Seven|Eight|Nine~|One|Two|Three|Four|Five|Six|Seven|Eight|Nine~Ten|Eleven|Twelve|Thirteen|Fourteen|Fifteen|Sixteen|Seventeen|Eighteen|Nineteen~||Twenty|Thirty|Forty|Fifty|Sixty|Seventy|Eighty|Ninety~|One hundred|Two hundred|Three hundred|Four hundred|Five hundred|Six hundred|Seven hundred|Eight hundred|Nine hundred~Thousand|Thousand|Thousand|0~Million|Million|Million|0~Billion|Billion|Billion|0~Zero"
    ),

    comment_currency, "--- 2. CURRENCY DATA ---",
    RAW_CURR_STR, SWITCH(target_lang,
        "UA", "UAH|гривня|гривні|гривень|1|копійка|копійки|копійок|1~USD|долар|долари|доларів|0|цент|цента|центів|0~EUR|євро|євро|євро|0|євроцент|євроцента|євроцентів|0~JPY|єна|єни|єн|1|сен|сени|сенів|0~GBP|фунт|фунти|фунтів|0|пенні|пенні|пенні|0~CHF|франк|франки|франків|0|сантим|сантими|сантимів|0~CAD|долар|долари|доларів|0|цент|цента|центів|0~AUD|долар|долари|доларів|0|цент|цента|центів|0~NZD|долар|долари|доларів|0|цент|цента|центів|0~CNY|юань|юані|юанів|0|фень|фені|фенів|0~SGD|долар|долари|доларів|0|цент|цента|центів|0~HKD|долар|долари|доларів|0|цент|цента|центів|0~ZAR|ренд|ренди|рендів|0|цент|цента|центів|0~SEK|крона|крони|крон|1|ере|ере|ере|0~NOK|крона|крони|крон|1|ере|ере|ере|0~MXN|песо|песо|песо|0|сентаво|сентаво|сентаво|0",
        
        "RU", "UAH|гривна|гривны|гривен|1|копейка|копейки|копеек|1~USD|доллар|доллара|долларов|0|цент|цента|центов|0~EUR|евро|евро|евро|0|евроцент|евроцента|евроцентов|0~JPY|иена|иены|иен|1|сен|сена|сенов|0~GBP|фунт|фунта|фунтов|0|пенни|пенни|пенни|0~CHF|франк|франка|франков|0|сантим|сантима|сантимов|0~CAD|доллар|доллара|долларов|0|цент|цента|центов|0~AUD|доллар|доллара|долларов|0|цент|цента|центов|0~NZD|доллар|доллара|долларов|0|цент|цента|центов|0~CNY|юань|юаня|юаней|0|фэнь|фэня|фэней|0~SGD|доллар|доллара|долларов|0|цент|цента|центов|0~HKD|доллар|доллара|долларов|0|цент|цента|центов|0~ZAR|рэнд|рэнда|рэндов|0|цент|цента|центов|0~SEK|крона|кроны|крон|1|эре|эре|эре|0~NOK|крона|кроны|крон|1|эре|эре|эре|0~MXN|песо|песо|песо|0|сентаво|сентаво|сентаво|0",
        
        "DE", "UAH|Hrywnja|Hrywnja|Hrywnja|1|Kopeke|Kopeken|Kopeken|1~USD|Dollar|Dollar|Dollar|0|Cent|Cent|Cent|0~EUR|Euro|Euro|Euro|0|Cent|Cent|Cent|0~JPY|Yen|Yen|Yen|0|Sen|Sen|Sen|0~GBP|Pfund|Pfund|Pfund|0|Penny|Pence|Pence|0~CHF|Franken|Franken|Franken|0|Rappen|Rappen|Rappen|0~CAD|Dollar|Dollar|Dollar|0|Cent|Cent|Cent|0~AUD|Dollar|Dollar|Dollar|0|Cent|Cent|Cent|0~NZD|Dollar|Dollar|Dollar|0|Cent|Cent|Cent|0~CNY|Yuan|Yuan|Yuan|0|Fen|Fen|Fen|0~SGD|Dollar|Dollar|Dollar|0|Cent|Cent|Cent|0~HKD|Dollar|Dollar|Dollar|0|Cent|Cent|Cent|0~ZAR|Rand|Rand|Rand|0|Cent|Cent|Cent|0~SEK|Krone|Kronen|Kronen|1|Öre|Öre|Öre|0~NOK|Krone|Kronen|Kronen|1|Öre|Öre|Öre|0~MXN|Peso|Pesos|Pesos|0|Centavo|Centavos|Centavos|0",
        
        "ES", "UAH|Grivna|Grivnas|Grivnas|1|kopek|kopeks|kopeks|0~USD|Dólar|Dólares|Dólares|0|centavo|centavos|centavos|0~EUR|Euro|Euros|Euros|0|céntimo|céntimos|céntimos|0~JPY|Yen|Yenes|Yenes|0|sen|sen|sen|0~GBP|Libra|Libras|Libras|1|penique|peniques|peniques|0~CHF|Franco|Francos|Francos|0|céntimo|céntimos|céntimos|0~CAD|Dólar|Dólares|Dólares|0|centavo|centavos|centavos|0~AUD|Dólar|Dólares|Dólares|0|centavo|centavos|centavos|0~NZD|Dólar|Dólares|Dólares|0|centavo|centavos|centavos|0~CNY|Yuan|Yuanes|Yuanes|0|fen|fen|fen|0~SGD|Dólar|Dólares|Dólares|0|centavo|centavos|centavos|0~HKD|Dólar|Dólares|Dólares|0|centavo|centavos|centavos|0~ZAR|Rand|Rands|Rands|0|centavo|centavos|centavos|0~SEK|Corona|Coronas|Coronas|1|öre|öre|öre|0~NOK|Corona|Coronas|Coronas|1|öre|öre|öre|0~MXN|Peso|Pesos|Pesos|0|centavo|centavos|centavos|0",
        
        "UAH|Hryvnia|Hryvnias|Hryvnias|0|kopeck|kopecks|kopecks|0~USD|Dollar|Dollars|Dollars|0|cent|cents|cents|0~EUR|Euro|Euros|Euros|0|cent|cents|cents|0~JPY|Yen|Yen|Yen|0|sen|sen|sen|0~GBP|Pound|Pounds|Pounds|0|penny|pence|pence|0~CHF|Franc|Francs|Francs|0|centime|centimes|centimes|0~CAD|Dollar|Dollars|Dollars|0|cent|cents|cents|0~AUD|Dollar|Dollars|Dollars|0|cent|cents|cents|0~NZD|Dollar|Dollars|Dollars|0|cent|cents|cents|0~CNY|Yuan|Yuan|Yuan|0|fen|fen|fen|0~SGD|Dollar|Dollars|Dollars|0|cent|cents|cents|0~HKD|Dollar|Dollars|Dollars|0|cent|cents|cents|0~ZAR|Rand|Rand|Rand|0|cent|cents|cents|0~SEK|Krona|Kronor|Kronor|0|ore|ore|ore|0~NOK|Krone|Kroner|Kroner|0|ore|ore|ore|0~MXN|Peso|Pesos|Pesos|0|centavo|centavos|centavos|0"
    ),

    comment_parsing, "--- 3. PARSING DATA ARRAYS ---",
    NUM_GROUPS, TOCOL(SPLIT(RAW_NUM_STR, "~")),
    
    Words_OnesM, TOCOL(SPLIT(INDEX(NUM_GROUPS, 1), "|", FALSE, FALSE)),
    Words_OnesF, TOCOL(SPLIT(INDEX(NUM_GROUPS, 2), "|", FALSE, FALSE)),
    Words_Teens, TOCOL(SPLIT(INDEX(NUM_GROUPS, 3), "|", FALSE, FALSE)),
    Words_Tens,  TOCOL(SPLIT(INDEX(NUM_GROUPS, 4), "|", FALSE, FALSE)),
    Words_Hund,  TOCOL(SPLIT(INDEX(NUM_GROUPS, 5), "|", FALSE, FALSE)),
    Words_Thous, TOCOL(SPLIT(INDEX(NUM_GROUPS, 6), "|", FALSE, FALSE)),
    Words_Mill,  TOCOL(SPLIT(INDEX(NUM_GROUPS, 7), "|", FALSE, FALSE)),
    Words_Bill,  TOCOL(SPLIT(INDEX(NUM_GROUPS, 8), "|", FALSE, FALSE)),
    Words_Dummy, TOCOL(SPLIT("|||0", "|", FALSE, FALSE)),
    Word_Zero,   INDEX(NUM_GROUPS, 9),

    CURR_LIST, TOCOL(SPLIT(RAW_CURR_STR, "~")),
    CURR_FOUND, XLOOKUP("*"&curr_code&"*", CURR_LIST, CURR_LIST, INDEX(CURR_LIST,1), 2),
    DATA_CURR, TOCOL(SPLIT(CURR_FOUND, "|")),
    CurrIsFem, INDEX(DATA_CURR, 5),
    CentIsFem, INDEX(DATA_CURR, 9),

    comment_logic, "--- 4. CORE LOGIC ---",
    getDeclension, LAMBDA(n, 
        LET(nn, MOD(ABS(n),100), r, MOD(nn,10),
            IF(OR(target_lang="US", target_lang="UK", target_lang="EN", target_lang="DE", target_lang="ES"), IF(n=1,1,2),
            IF(AND(nn>=11, nn<=19), 3, IF(r=1, 1, IF(AND(r>=2, r<=4), 2, 3))))
        )
    ),

    getTriad, LAMBDA(n, isFem, isThousand,
        LET(
            h, INT(n/100), t, MOD(n,100), d, MOD(n,10),
            txtH, IF(h=0, "", IF(AND(target_lang="ES", h=1, t=0), "cien", INDEX(Words_Hund, h+1))),
            txtTensPart, 
            IF(AND(t>=10, t<=19), INDEX(Words_Teens, t-9),
                IF(t>=20, 
                    LET(
                        tenW, INDEX(Words_Tens, INT(t/10)+1),
                        oneW, IF(d=0, "", INDEX(IF(isFem, Words_OnesF, Words_OnesM), d+1)),
                        IF(target_lang="DE", 
                            IF(d=0, tenW, IF(d=1,"ein",oneW) & "und" & tenW),
                            IF(target_lang="ES", 
                                IF(d=0, tenW, tenW & " y " & IF(AND(isThousand, d=1), "", oneW)),
                                tenW & " " & oneW 
                            )
                        )
                    ),
                    IF(d=0, "", 
                       IF(AND(target_lang="ES", isThousand, d=1), "", INDEX(IF(isFem, Words_OnesF, Words_OnesM), d+1))
                    )
                )
            ),
            TRIM(txtH & " " & txtTensPart)
        )
    ),

    processPeriod, LAMBDA(v, idx, acc,
        LET(
            currTriad, MOD(INT(v / 10^(3*(idx-1))), 1000),
            periodData, CHOOSE(idx, Words_Dummy, Words_Thous, Words_Mill, Words_Bill),
            isFemPeriod, IF(idx=1, CurrIsFem, IF(idx>1, INDEX(periodData, 4), 0)),
            IF(currTriad = 0, acc,
                LET(
                    triadWords, getTriad(currTriad, isFemPeriod, (idx=2)),
                    decl, getDeclension(currTriad),
                    pName, INDEX(periodData, decl),
                    TRIM(triadWords & " " & pName & " " & acc)
                )
            )
        )
    ),

    calculateText, LAMBDA(n,
        LET(
            valRound, ROUND(n, 2),
            valInt, INT(valRound),
            valCents, ROUND((valRound - valInt) * 100, 0),
            
            strInt, IF(valInt=0, Word_Zero, REDUCE("", {1;2;3;4}, LAMBDA(acc, i, processPeriod(valInt, i, acc)))),
            currWord, INDEX(DATA_CURR, getDeclension(valInt) + 1),
            
            strCents, IF(valCents=0, "", getTriad(valCents, CentIsFem, FALSE)),
            centWord, IF(valCents=0, "", INDEX(DATA_CURR, getDeclension(valCents) + 5)),
            
            fullPhrase, TRIM(strInt & " " & currWord & " " & strCents & " " & centWord),
            UPPER(LEFT(fullPhrase,1)) & MID(fullPhrase, 2, 999)
        )
    ),

    IF(NOT(ISNUMBER(val)), "", IF(val<0, "Error: Negative", calculateText(val)))
)
```
</details>

Now use it like this:
`=AMOUNT_TO_WORDS(1250.50, "USD", "EN")`
// Output: One Thousand Two hundred Fifty Dollars Fifty Cents

### Method 2: Direct Paste (No Named Function)

If you cannot or do not want to use Named Functions, you can paste the formula directly into a cell.

1.  Copy the code from the **Source Code** section above.
2.  Paste it into your target cell (e.g., `D2`).
3.  **Crucial Step:** You must manually replace the variable names at the very top of the formula with your actual cell references.

**Change this:**
```excel
=LET(
    val, val,
    curr_code, UPPER(curr_code),
    target_lang, UPPER(target_lang),
    ...
```
**To this (example):**
```excel
=LET(
    val, A2,                 <-- Reference to your Number
    curr_code, UPPER(B2),    <-- Reference to Currency Code
    target_lang, UPPER(C2),  <-- Reference to Language Code
    ...
```

## 📚 Supported Lists

### Languages

| Code | Language | Native Name |
| :--- | :--- | :--- |
| **EN** | English (US/UK) | English |
| **UA** | Ukrainian | Українська |
| **RU** | Russian | Русский |
| **DE** | German | Deutsch |
| **ES** | Spanish | Español |

### Currencies

| Code | Currency |
| :--- | :--- |
| **USD** | US Dollar |
| **EUR** | Euro |
| **UAH** | Ukrainian Hryvnia |
| **GBP** | Pound Sterling |
| **JPY** | Japanese Yen |
| **CHF** | Swiss Franc |
| **CNY** | Chinese Yuan |
| **CAD** | Canadian Dollar |
| **AUD** | Australian Dollar |
| **NZD** | New Zealand Dollar |
| **SGD** | Singapore Dollar |
| **HKD** | Hong Kong Dollar |
| **ZAR** | South African Rand |
| **SEK** | Swedish Krona |
| **NOK** | Norwegian Krone |
| **MXN** | Mexican Peso |

## ⚠️ Technical Notes & Limits

1.  **Max Value:** `999,999,999,999.99` (Up to 999 Billions). Numbers ≥ 1 Trillion are not supported in this version.
2.  **Rounding:** Automatically rounds numbers to **2 decimal places** (standard financial rounding).
3.  **Case Sensitivity:** Input codes are case-insensitive (`usd`, `USD`, `uSd` work equally well).
4.  **Compatibility:** Requires **Google Sheets** or **Excel 365/2021+**. Will not work in older Excel versions that lack `LET` and `LAMBDA` functions.
5.  **Grammar:** Outputs text in the **Nominative case** (Standard for invoices, contracts, and checks).
6.  **Negative Numbers:** Currently outputs an error string `"Error: Negative"`.

## 🛠️ Customizing & Contributing

Missing your language? You can add it by modifying two variables in the formula:

1.  **`RAW_NUM_STR`**: Add your language code and the corresponding string of numerals (ones, teens, tens, hundreds, magnitudes) separated by `~`.
2.  **`RAW_CURR_STR`**: Add your currency definitions in the format `CODE|Main1|Main2|Main5|Gender|Sub1|Sub2|Sub5|GenderSub`.

### 💡 Pro Tip: Use AI
Instead of formatting these strings manually, I highly recommend using **Gemini 3 Pro**. It handles the array structures and linguistic logic perfectly.

**Prompt example:** > *"Here is a Google Sheets LET formula. Please add French language support to the RAW_NUM_STR variable and CHF to RAW_CURR_STR following the existing pattern."*

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
