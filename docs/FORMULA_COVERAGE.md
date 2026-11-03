# Formula coverage

_Generated from the function catalog (`src/core/engine/catalog.json`, built from Microsoft’s “Excel functions (alphabetical)” page) and the function tests (`tests/unit/functions/cases.ts`). Do not edit by hand: run `UPDATE_DOCS=1 npx vitest run tests/unit/docs.test.ts`._

**499 of 526** Excel worksheet functions are calculated. **481** of them have dedicated test cases (973 formulas, most taken from the examples in Microsoft’s documentation).
The formulas are also run through LibreOffice Calc as an independent check ([LIBREOFFICE_CROSSCHECK.md](LIBREOFFICE_CROSSCHECK.md)). Simple Cell has not been checked against Excel itself: no copy of Excel was available to the project.

A formula that uses a function Simple Cell does not calculate is kept unchanged and shows the result saved in the file, marked as not recalculated. Without a saved result it shows `#NAME?`. It is never shown as 0.

| Status | Meaning |
|---|---|
| ✅ | Calculated, with test cases |
| 🟡 | Calculated, no dedicated test case yet |
| ⛔ | Not calculated (formula kept; saved result shown) |

## By category

| Category | Calculated | Tested | Total |
|---|---|---|---|
| AI | 0 | 0 | 1 |
| Add-in and Automation | 1 | 1 | 3 |
| Compatibility | 40 | 32 | 40 |
| Cube | 0 | 0 | 7 |
| Database | 12 | 12 | 12 |
| Date and time | 25 | 25 | 25 |
| Engineering | 54 | 54 | 54 |
| Financial | 55 | 55 | 55 |
| Information | 21 | 21 | 22 |
| Logical | 19 | 19 | 19 |
| Lookup and reference | 36 | 36 | 41 |
| Math and trigonometry | 80 | 80 | 80 |
| Python | 0 | 0 | 1 |
| Statistical | 107 | 106 | 111 |
| Text | 48 | 39 | 50 |
| Web | 1 | 1 | 5 |
| **All** | **499** | **481** | **526** |

## Known deviations from Excel

| Function | Deviation |
|---|---|
| DURATION | Simple Cell computes the exact Macaulay duration over coupon periods; Excel’s documented examples differ by up to 2×10⁻⁶ (relative). |
| MDURATION | Derived from DURATION (same deviation). |

## AI

| Function | Status | Since | Notes |
|---|---|---|---|
| COPILOT | ⛔ | Microsoft 365 | Microsoft Copilot service; not available. |

## Add-in and Automation

| Function | Status | Since | Notes |
|---|---|---|---|
| CALL | ⛔ | — | Calls DLL procedures; not available (security). |
| EUROCONVERT | ✅ | — | Fixed euro conversion rates for the 20 legacy currencies. |
| REGISTER.ID | ⛔ | — | Registers DLL procedures; not available (security). |

## Compatibility

| Function | Status | Since | Notes |
|---|---|---|---|
| BETADIST | ✅ | — |  |
| BETAINV | 🟡 | — |  |
| BINOMDIST | ✅ | — |  |
| CEILING | ✅ | — |  |
| CHIDIST | ✅ | — |  |
| CHIINV | ✅ | — |  |
| CHITEST | 🟡 | — |  |
| CONFIDENCE | ✅ | — |  |
| COVAR | ✅ | — |  |
| CRITBINOM | ✅ | — |  |
| EXPONDIST | ✅ | — |  |
| FDIST | ✅ | — |  |
| FINV | ✅ | — |  |
| FLOOR | ✅ | — |  |
| FTEST | 🟡 | — |  |
| GAMMADIST | ✅ | — |  |
| GAMMAINV | 🟡 | — |  |
| HYPGEOMDIST | ✅ | — |  |
| LOGINV | 🟡 | — |  |
| LOGNORMDIST | ✅ | — |  |
| MODE | ✅ | — |  |
| NEGBINOMDIST | ✅ | — |  |
| NORM.INV | ✅ | 2010 |  |
| NORMDIST | ✅ | — |  |
| NORMSDIST | ✅ | — |  |
| NORMSINV | ✅ | — |  |
| PERCENTILE | ✅ | — |  |
| PERCENTRANK | 🟡 | — |  |
| POISSON | ✅ | — |  |
| QUARTILE | ✅ | — |  |
| RANK | ✅ | — |  |
| STDEV | ✅ | — |  |
| STDEVP | ✅ | — |  |
| TDIST | ✅ | — |  |
| TINV | ✅ | — |  |
| TTEST | 🟡 | — |  |
| VAR | ✅ | — |  |
| VARP | ✅ | — |  |
| WEIBULL | ✅ | — |  |
| ZTEST | 🟡 | — |  |

## Cube

| Function | Status | Since | Notes |
|---|---|---|---|
| CUBEKPIMEMBER | ⛔ | — | OLAP cube connections are not supported. |
| CUBEMEMBER | ⛔ | — | OLAP cube connections are not supported. |
| CUBEMEMBERPROPERTY | ⛔ | — | OLAP cube connections are not supported. |
| CUBERANKEDMEMBER | ⛔ | — | OLAP cube connections are not supported. |
| CUBESET | ⛔ | — | OLAP cube connections are not supported. |
| CUBESETCOUNT | ⛔ | — | OLAP cube connections are not supported. |
| CUBEVALUE | ⛔ | — | OLAP cube connections are not supported. |

## Database

| Function | Status | Since | Notes |
|---|---|---|---|
| DAVERAGE | ✅ | — | Criteria ranges with formula (computed) criteria are not supported yet; value criteria work as in Excel. |
| DCOUNT | ✅ | — |  |
| DCOUNTA | ✅ | — |  |
| DGET | ✅ | — |  |
| DMAX | ✅ | — |  |
| DMIN | ✅ | — |  |
| DPRODUCT | ✅ | — |  |
| DSTDEV | ✅ | — |  |
| DSTDEVP | ✅ | — |  |
| DSUM | ✅ | — |  |
| DVAR | ✅ | — |  |
| DVARP | ✅ | — |  |

## Date and time

| Function | Status | Since | Notes |
|---|---|---|---|
| DATE | ✅ | — |  |
| DATEDIF | ✅ | — |  |
| DATEVALUE | ✅ | — |  |
| DAY | ✅ | — |  |
| DAYS | ✅ | 2013 |  |
| DAYS360 | ✅ | — |  |
| EDATE | ✅ | — |  |
| EOMONTH | ✅ | — |  |
| HOUR | ✅ | — |  |
| ISOWEEKNUM | ✅ | 2013 |  |
| MINUTE | ✅ | — |  |
| MONTH | ✅ | — |  |
| NETWORKDAYS | ✅ | — |  |
| NETWORKDAYS.INTL | ✅ | 2010 |  |
| NOW | ✅ | — |  |
| SECOND | ✅ | — |  |
| TIME | ✅ | — |  |
| TIMEVALUE | ✅ | — |  |
| TODAY | ✅ | — |  |
| WEEKDAY | ✅ | — |  |
| WEEKNUM | ✅ | — |  |
| WORKDAY | ✅ | — |  |
| WORKDAY.INTL | ✅ | 2010 |  |
| YEAR | ✅ | — |  |
| YEARFRAC | ✅ | — |  |

## Engineering

| Function | Status | Since | Notes |
|---|---|---|---|
| BESSELI | ✅ | — |  |
| BESSELJ | ✅ | — |  |
| BESSELK | ✅ | — |  |
| BESSELY | ✅ | — |  |
| BIN2DEC | ✅ | — |  |
| BIN2HEX | ✅ | — |  |
| BIN2OCT | ✅ | — |  |
| BITAND | ✅ | 2013 |  |
| BITLSHIFT | ✅ | 2013 |  |
| BITOR | ✅ | 2013 |  |
| BITRSHIFT | ✅ | 2013 |  |
| BITXOR | ✅ | 2013 |  |
| COMPLEX | ✅ | — |  |
| CONVERT | ✅ | — |  |
| DEC2BIN | ✅ | — |  |
| DEC2HEX | ✅ | — |  |
| DEC2OCT | ✅ | — |  |
| DELTA | ✅ | — |  |
| ERF | ✅ | — |  |
| ERF.PRECISE | ✅ | 2010 |  |
| ERFC | ✅ | — |  |
| ERFC.PRECISE | ✅ | 2010 |  |
| GESTEP | ✅ | — |  |
| HEX2BIN | ✅ | — |  |
| HEX2DEC | ✅ | — |  |
| HEX2OCT | ✅ | — |  |
| IMABS | ✅ | — |  |
| IMAGINARY | ✅ | — |  |
| IMARGUMENT | ✅ | — |  |
| IMCONJUGATE | ✅ | — |  |
| IMCOS | ✅ | — |  |
| IMCOSH | ✅ | 2013 |  |
| IMCOT | ✅ | 2013 |  |
| IMCSC | ✅ | 2013 |  |
| IMCSCH | ✅ | 2013 |  |
| IMDIV | ✅ | — |  |
| IMEXP | ✅ | — |  |
| IMLN | ✅ | — |  |
| IMLOG10 | ✅ | — |  |
| IMLOG2 | ✅ | — |  |
| IMPOWER | ✅ | — |  |
| IMPRODUCT | ✅ | — |  |
| IMREAL | ✅ | — |  |
| IMSEC | ✅ | 2013 |  |
| IMSECH | ✅ | 2013 |  |
| IMSIN | ✅ | — |  |
| IMSINH | ✅ | 2013 |  |
| IMSQRT | ✅ | — |  |
| IMSUB | ✅ | — |  |
| IMSUM | ✅ | — |  |
| IMTAN | ✅ | 2013 |  |
| OCT2BIN | ✅ | — |  |
| OCT2DEC | ✅ | — |  |
| OCT2HEX | ✅ | — |  |

## Financial

| Function | Status | Since | Notes |
|---|---|---|---|
| ACCRINT | ✅ | — | Accrued interest from issue to settlement (calc_method TRUE); first_interest is not used for the FALSE method yet. |
| ACCRINTM | ✅ | — |  |
| AMORDEGRC | ✅ | — |  |
| AMORLINC | ✅ | — |  |
| COUPDAYBS | ✅ | — |  |
| COUPDAYS | ✅ | — |  |
| COUPDAYSNC | ✅ | — |  |
| COUPNCD | ✅ | — |  |
| COUPNUM | ✅ | — |  |
| COUPPCD | ✅ | — |  |
| CUMIPMT | ✅ | — |  |
| CUMPRINC | ✅ | — |  |
| DB | ✅ | — |  |
| DDB | ✅ | — |  |
| DISC | ✅ | — |  |
| DOLLARDE | ✅ | — |  |
| DOLLARFR | ✅ | — |  |
| DURATION | ✅ | — |  |
| EFFECT | ✅ | — |  |
| FV | ✅ | — |  |
| FVSCHEDULE | ✅ | — |  |
| INTRATE | ✅ | — |  |
| IPMT | ✅ | — |  |
| IRR | ✅ | — |  |
| ISPMT | ✅ | — |  |
| MDURATION | ✅ | — |  |
| MIRR | ✅ | — |  |
| NOMINAL | ✅ | — |  |
| NPER | ✅ | — |  |
| NPV | ✅ | — |  |
| ODDFPRICE | ✅ | — | Short first periods follow Microsoft’s formula exactly; long first periods use the documented quasi-coupon method (not verified against Excel). |
| ODDFYIELD | ✅ | — | Solves ODDFPRICE for the yield. |
| ODDLPRICE | ✅ | — |  |
| ODDLYIELD | ✅ | — |  |
| PDURATION | ✅ | 2013 |  |
| PMT | ✅ | — |  |
| PPMT | ✅ | — |  |
| PRICE | ✅ | — |  |
| PRICEDISC | ✅ | — |  |
| PRICEMAT | ✅ | — |  |
| PV | ✅ | — |  |
| RATE | ✅ | — |  |
| RECEIVED | ✅ | — |  |
| RRI | ✅ | 2013 |  |
| SLN | ✅ | — |  |
| SYD | ✅ | — |  |
| TBILLEQ | ✅ | — |  |
| TBILLPRICE | ✅ | — |  |
| TBILLYIELD | ✅ | — |  |
| VDB | ✅ | — |  |
| XIRR | ✅ | — |  |
| XNPV | ✅ | — |  |
| YIELD | ✅ | — |  |
| YIELDDISC | ✅ | — |  |
| YIELDMAT | ✅ | — |  |

## Information

| Function | Status | Since | Notes |
|---|---|---|---|
| CELL | ✅ | — | "color", "prefix", "protect", "format" and "width" follow Excel’s rules for common formats; "filename" needs a saved file. |
| ERROR.TYPE | ✅ | — |  |
| INFO | ✅ | — | "osversion" and "system" report the platform Simple Cell runs on. |
| ISBLANK | ✅ | — |  |
| ISERR | ✅ | — |  |
| ISERROR | ✅ | — |  |
| ISEVEN | ✅ | — |  |
| ISFORMULA | ✅ | 2013 |  |
| ISLOGICAL | ✅ | — |  |
| ISNA | ✅ | — |  |
| ISNONTEXT | ✅ | — |  |
| ISNUMBER | ✅ | — |  |
| ISODD | ✅ | — |  |
| ISOMITTED | ✅ | 2024 |  |
| ISREF | ✅ | — |  |
| ISTEXT | ✅ | — |  |
| N | ✅ | — |  |
| NA | ✅ | — |  |
| SHEET | ✅ | 2013 |  |
| SHEETS | ✅ | 2013 |  |
| STOCKHISTORY | ⛔ | — | Microsoft online data service; not available. |
| TYPE | ✅ | — |  |

## Logical

| Function | Status | Since | Notes |
|---|---|---|---|
| AND | ✅ | — |  |
| BYCOL | ✅ | 2024 |  |
| BYROW | ✅ | 2024 |  |
| FALSE | ✅ | — |  |
| IF | ✅ | — |  |
| IFERROR | ✅ | — |  |
| IFNA | ✅ | 2013 |  |
| IFS | ✅ | 2019 |  |
| LAMBDA | ✅ | 2024 |  |
| LET | ✅ | 2021 |  |
| MAKEARRAY | ✅ | 2024 |  |
| MAP | ✅ | 2024 |  |
| NOT | ✅ | — |  |
| OR | ✅ | — |  |
| REDUCE | ✅ | 2024 |  |
| SCAN | ✅ | 2024 |  |
| SWITCH | ✅ | 2019 |  |
| TRUE | ✅ | — |  |
| XOR | ✅ | 2013 |  |

## Lookup and reference

| Function | Status | Since | Notes |
|---|---|---|---|
| ADDRESS | ✅ | — |  |
| AREAS | ✅ | — |  |
| CHOOSE | ✅ | — |  |
| CHOOSECOLS | ✅ | 2024 |  |
| CHOOSEROWS | ✅ | 2024 |  |
| COLUMN | ✅ | — |  |
| COLUMNS | ✅ | — |  |
| DROP | ✅ | 2024 |  |
| EXPAND | ✅ | 2024 |  |
| FIELDVALUE | ⛔ | Microsoft 365 | Linked data types (Stocks, Geography) are Microsoft online services; not available. |
| FILTER | ✅ | 2021 |  |
| FORMULATEXT | ✅ | 2013 |  |
| GETPIVOTDATA | ⛔ | — | Arrives with PivotTables (M6). |
| GROUPBY | ✅ | Microsoft 365 | Core grouping, aggregation, headers, totals and sorting; subtotal depth beyond 1 and field relationships are not implemented. |
| HLOOKUP | ✅ | — |  |
| HSTACK | ✅ | 2024 |  |
| HYPERLINK | ✅ | — |  |
| IMAGE | ⛔ | 2024 | Images in cells are planned. |
| INDEX | ✅ | — |  |
| INDIRECT | ✅ | — |  |
| LOOKUP | ✅ | — |  |
| MATCH | ✅ | — |  |
| OFFSET | ✅ | — |  |
| PIVOTBY | ⛔ | Microsoft 365 | Planned with PivotTables (M6). |
| ROW | ✅ | — |  |
| ROWS | ✅ | — |  |
| RTD | ⛔ | — | COM real-time data servers are Windows-only Excel add-ins; not available. |
| SORT | ✅ | 2021 |  |
| SORTBY | ✅ | 2021 |  |
| TAKE | ✅ | 2024 |  |
| TOCOL | ✅ | 2024 |  |
| TOROW | ✅ | 2024 |  |
| TRANSPOSE | ✅ | — |  |
| TRIMRANGE | ✅ | Microsoft 365 |  |
| UNIQUE | ✅ | 2021 |  |
| VLOOKUP | ✅ | — |  |
| VSTACK | ✅ | 2024 |  |
| WRAPCOLS | ✅ | 2024 |  |
| WRAPROWS | ✅ | 2024 |  |
| XLOOKUP | ✅ | 2021 |  |
| XMATCH | ✅ | 2021 |  |

## Math and trigonometry

| Function | Status | Since | Notes |
|---|---|---|---|
| ABS | ✅ | — |  |
| ACOS | ✅ | — |  |
| ACOSH | ✅ | — |  |
| ACOT | ✅ | 2013 |  |
| ACOTH | ✅ | 2013 |  |
| AGGREGATE | ✅ | — |  |
| ARABIC | ✅ | 2013 |  |
| ASIN | ✅ | — |  |
| ASINH | ✅ | — |  |
| ATAN | ✅ | — |  |
| ATAN2 | ✅ | — |  |
| ATANH | ✅ | — |  |
| BASE | ✅ | — |  |
| CEILING.MATH | ✅ | 2013 |  |
| CEILING.PRECISE | ✅ | — |  |
| COMBIN | ✅ | — |  |
| COMBINA | ✅ | 2013 |  |
| COS | ✅ | — |  |
| COSH | ✅ | — |  |
| COT | ✅ | 2013 |  |
| COTH | ✅ | 2013 |  |
| CSC | ✅ | 2013 |  |
| CSCH | ✅ | 2013 |  |
| DECIMAL | ✅ | 2013 |  |
| DEGREES | ✅ | — |  |
| EVEN | ✅ | — |  |
| EXP | ✅ | — |  |
| FACT | ✅ | — |  |
| FACTDOUBLE | ✅ | — |  |
| FLOOR.MATH | ✅ | 2013 |  |
| FLOOR.PRECISE | ✅ | — |  |
| GCD | ✅ | — |  |
| INT | ✅ | — |  |
| ISO.CEILING | ✅ | 2013 |  |
| LCM | ✅ | — |  |
| LN | ✅ | — |  |
| LOG | ✅ | — |  |
| LOG10 | ✅ | — |  |
| MDETERM | ✅ | — |  |
| MINVERSE | ✅ | — |  |
| MMULT | ✅ | — |  |
| MOD | ✅ | — |  |
| MROUND | ✅ | — |  |
| MULTINOMIAL | ✅ | — |  |
| MUNIT | ✅ | 2013 |  |
| ODD | ✅ | — |  |
| PERCENTOF | ✅ | Microsoft 365 |  |
| PI | ✅ | — |  |
| POWER | ✅ | — |  |
| PRODUCT | ✅ | — |  |
| QUOTIENT | ✅ | — |  |
| RADIANS | ✅ | — |  |
| RAND | ✅ | — |  |
| RANDARRAY | ✅ | 2021 |  |
| RANDBETWEEN | ✅ | — |  |
| ROMAN | ✅ | — |  |
| ROUND | ✅ | — |  |
| ROUNDDOWN | ✅ | — |  |
| ROUNDUP | ✅ | — |  |
| SEC | ✅ | 2013 |  |
| SECH | ✅ | 2013 |  |
| SEQUENCE | ✅ | 2021 |  |
| SERIESSUM | ✅ | — |  |
| SIGN | ✅ | — |  |
| SIN | ✅ | — |  |
| SINH | ✅ | — |  |
| SQRT | ✅ | — |  |
| SQRTPI | ✅ | — |  |
| SUBTOTAL | ✅ | — |  |
| SUM | ✅ | — |  |
| SUMIF | ✅ | — |  |
| SUMIFS | ✅ | — |  |
| SUMPRODUCT | ✅ | — |  |
| SUMSQ | ✅ | — |  |
| SUMX2MY2 | ✅ | — |  |
| SUMX2PY2 | ✅ | — |  |
| SUMXMY2 | ✅ | — |  |
| TAN | ✅ | — |  |
| TANH | ✅ | — |  |
| TRUNC | ✅ | — |  |

## Python

| Function | Status | Since | Notes |
|---|---|---|---|
| PY | ⛔ | Microsoft 365 | Python in Excel runs in Microsoft’s cloud; not available. |

## Statistical

| Function | Status | Since | Notes |
|---|---|---|---|
| AVEDEV | ✅ | — |  |
| AVERAGE | ✅ | — |  |
| AVERAGEA | ✅ | — |  |
| AVERAGEIF | ✅ | — |  |
| AVERAGEIFS | ✅ | — |  |
| BETA.DIST | ✅ | 2010 |  |
| BETA.INV | ✅ | 2010 |  |
| BINOM.DIST | ✅ | 2010 |  |
| BINOM.DIST.RANGE | ✅ | 2013 |  |
| BINOM.INV | ✅ | 2010 |  |
| CHISQ.DIST | ✅ | 2010 |  |
| CHISQ.DIST.RT | ✅ | 2010 |  |
| CHISQ.INV | ✅ | 2010 |  |
| CHISQ.INV.RT | ✅ | 2010 |  |
| CHISQ.TEST | ✅ | 2010 |  |
| CONFIDENCE.NORM | ✅ | 2010 |  |
| CONFIDENCE.T | ✅ | 2010 |  |
| CORREL | ✅ | — |  |
| COUNT | ✅ | — |  |
| COUNTA | ✅ | — |  |
| COUNTBLANK | ✅ | — |  |
| COUNTIF | ✅ | — |  |
| COUNTIFS | ✅ | — |  |
| COVARIANCE.P | ✅ | 2010 |  |
| COVARIANCE.S | ✅ | 2010 |  |
| DEVSQ | ✅ | — |  |
| EXPON.DIST | ✅ | 2010 |  |
| F.DIST | ✅ | 2010 |  |
| F.DIST.RT | ✅ | 2010 |  |
| F.INV | ✅ | 2010 |  |
| F.INV.RT | ✅ | 2010 |  |
| F.TEST | ✅ | 2010 |  |
| FISHER | ✅ | — |  |
| FISHERINV | ✅ | — |  |
| FORECAST | ✅ | — |  |
| FORECAST.ETS | ⛔ | 2016 | Exponential triple smoothing (AAA ETS) is not implemented yet; formulas keep the value saved in the file. |
| FORECAST.ETS.CONFINT | ⛔ | 2016 | See FORECAST.ETS. |
| FORECAST.ETS.SEASONALITY | ⛔ | 2016 | See FORECAST.ETS. |
| FORECAST.ETS.STAT | ⛔ | 2016 | See FORECAST.ETS. |
| FORECAST.LINEAR | ✅ | 2016 |  |
| FREQUENCY | ✅ | — |  |
| GAMMA | ✅ | 2013 |  |
| GAMMA.DIST | ✅ | 2010 |  |
| GAMMA.INV | ✅ | 2010 |  |
| GAMMALN | ✅ | — |  |
| GAMMALN.PRECISE | ✅ | 2010 |  |
| GAUSS | ✅ | 2013 |  |
| GEOMEAN | ✅ | — |  |
| GROWTH | ✅ | — |  |
| HARMEAN | ✅ | — |  |
| HYPGEOM.DIST | ✅ | — |  |
| INTERCEPT | ✅ | — |  |
| KURT | ✅ | — |  |
| LARGE | ✅ | — |  |
| LINEST | ✅ | — |  |
| LOGEST | ✅ | — |  |
| LOGNORM.DIST | ✅ | 2010 |  |
| LOGNORM.INV | ✅ | 2010 |  |
| MAX | ✅ | — |  |
| MAXA | ✅ | — |  |
| MAXIFS | ✅ | 2019 |  |
| MEDIAN | ✅ | — |  |
| MIN | ✅ | — |  |
| MINA | ✅ | — |  |
| MINIFS | ✅ | 2019 |  |
| MODE.MULT | ✅ | 2010 |  |
| MODE.SNGL | ✅ | 2010 |  |
| NEGBINOM.DIST | ✅ | 2010 |  |
| NORM.DIST | ✅ | 2010 |  |
| NORM.S.DIST | ✅ | 2010 |  |
| NORM.S.INV | ✅ | 2010 |  |
| NORMINV | 🟡 | — |  |
| PEARSON | ✅ | — |  |
| PERCENTILE.EXC | ✅ | 2010 |  |
| PERCENTILE.INC | ✅ | 2010 |  |
| PERCENTRANK.EXC | ✅ | 2010 |  |
| PERCENTRANK.INC | ✅ | 2010 |  |
| PERMUT | ✅ | — |  |
| PERMUTATIONA | ✅ | 2013 |  |
| PHI | ✅ | 2013 |  |
| POISSON.DIST | ✅ | 2010 |  |
| PROB | ✅ | — |  |
| QUARTILE.EXC | ✅ | 2010 |  |
| QUARTILE.INC | ✅ | 2010 |  |
| RANK.AVG | ✅ | 2010 |  |
| RANK.EQ | ✅ | 2010 |  |
| RSQ | ✅ | — |  |
| SKEW | ✅ | — |  |
| SKEW.P | ✅ | 2013 |  |
| SLOPE | ✅ | — |  |
| SMALL | ✅ | — |  |
| STANDARDIZE | ✅ | — |  |
| STDEV.P | ✅ | 2010 |  |
| STDEV.S | ✅ | 2010 |  |
| STDEVA | ✅ | — |  |
| STDEVPA | ✅ | — |  |
| STEYX | ✅ | — |  |
| T.DIST | ✅ | 2010 |  |
| T.DIST.2T | ✅ | 2010 |  |
| T.DIST.RT | ✅ | 2010 |  |
| T.INV | ✅ | 2010 |  |
| T.INV.2T | ✅ | 2010 |  |
| T.TEST | ✅ | 2010 |  |
| TREND | ✅ | — |  |
| TRIMMEAN | ✅ | — |  |
| VAR.P | ✅ | 2010 |  |
| VAR.S | ✅ | 2010 |  |
| VARA | ✅ | — |  |
| VARPA | ✅ | — |  |
| WEIBULL.DIST | ✅ | 2010 |  |
| Z.TEST | ✅ | 2010 |  |

## Text

| Function | Status | Since | Notes |
|---|---|---|---|
| ARRAYTOTEXT | ✅ | 2021 |  |
| ASC | ✅ | — | Full-width to half-width ASCII and spaces; katakana conversion is not included. |
| BAHTTEXT | ✅ | — |  |
| CHAR | ✅ | — |  |
| CLEAN | ✅ | — |  |
| CODE | ✅ | — |  |
| CONCAT | ✅ | 2019 |  |
| CONCATENATE | ✅ | — |  |
| DBCS | ✅ | 2013 | Half-width to full-width ASCII and spaces; katakana conversion is not included. |
| DETECTLANGUAGE | ⛔ | Microsoft 365 | Microsoft online service; not available. |
| DOLLAR | ✅ | — |  |
| EXACT | ✅ | — |  |
| FIND | ✅ | — |  |
| FINDB | 🟡 | — | Same as FIND (single-byte languages). |
| FIXED | ✅ | — |  |
| JIS | 🟡 | — | Same as DBCS. |
| LEFT | ✅ | — |  |
| LEFTB | 🟡 | — | Same as LEFT (single-byte languages). |
| LEN | ✅ | — |  |
| LENB | 🟡 | — | Counts characters (single-byte languages); double-byte byte counts are not used. |
| LOWER | ✅ | — |  |
| MID | ✅ | — |  |
| MIDB | 🟡 | — | Same as MID (single-byte languages). |
| NUMBERVALUE | ✅ | 2013 |  |
| PHONETIC | 🟡 | — | Returns the text itself: furigana (phonetic guides) are not stored. |
| PROPER | ✅ | — |  |
| REGEXEXTRACT | ✅ | Microsoft 365 | Uses JavaScript regular expressions, which differ from Excel’s PCRE2 in rare constructs. |
| REGEXREPLACE | ✅ | Microsoft 365 | See REGEXEXTRACT. |
| REGEXTEST | ✅ | Microsoft 365 | See REGEXEXTRACT. |
| REPLACE | ✅ | — |  |
| REPLACEB | 🟡 | — | Same as REPLACE (single-byte languages). |
| REPT | ✅ | — |  |
| RIGHT | ✅ | — |  |
| RIGHTB | 🟡 | — | Same as RIGHT (single-byte languages). |
| SEARCH | ✅ | — |  |
| SEARCHB | 🟡 | — | Same as SEARCH (single-byte languages). |
| SUBSTITUTE | ✅ | — |  |
| T | ✅ | — |  |
| TEXT | ✅ | — |  |
| TEXTAFTER | ✅ | 2024 |  |
| TEXTBEFORE | ✅ | 2024 |  |
| TEXTJOIN | ✅ | 2019 |  |
| TEXTSPLIT | ✅ | 2024 |  |
| TRANSLATE | ⛔ | Microsoft 365 | Microsoft online translation service; not available. |
| TRIM | ✅ | — |  |
| UNICHAR | ✅ | 2013 |  |
| UNICODE | ✅ | 2013 |  |
| UPPER | ✅ | — |  |
| VALUE | ✅ | — |  |
| VALUETOTEXT | ✅ | 2021 |  |

## Web

| Function | Status | Since | Notes |
|---|---|---|---|
| ENCODEURL | ✅ | 2013 |  |
| FILTERXML | ⛔ | 2013 | Needs an XPath engine; planned. |
| IMPORTCSV | ⛔ | Microsoft 365 | Needs file/network access during calculation; planned. |
| IMPORTTEXT | ⛔ | Microsoft 365 | Needs file/network access during calculation; planned. |
| WEBSERVICE | ⛔ | 2013 | Needs network access during calculation; planned behind an explicit permission. |
