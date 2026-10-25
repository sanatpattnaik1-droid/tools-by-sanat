# Ribbon button map (DRAFT, confirm before building customUI14.xml)

Tab: **Tools By Sanat**. Status: OK = name matches a Sub exactly, ? = my guess, please confirm.

## Group: HA
| Button label | Macro (Module.Sub) | Status |
|---|---|---|
| Cmp Wise Sale Register (Detail) | COMP_W_REG_D.COMP_W_REG_D | OK |
| DABUR E-INVOICE | DABUR.DABUR_E_INVOICE | OK |
| IRN Upload Format | IRN_UPLOAD.IRN_UPLOAD | OK |
| Collection Outstanding Report Setup | Collection_Outstanding_Report1.Collection_Outstanding_Report | OK |
| Van_Wise Outstanding Report | Van_Wise_Outstanding_Report.Van_Wise_Outstanding_Report | OK |
| Keep Only Highlighted Cell Rows | Keep_Only_Selected.Keep_Only_Selected | ? (label vs name) |
| Fill Down Blanks | Fill_Down.FillDownBlanks | OK |
| Party Wise Receiving Summary | Sale_Receiving_Summary.Receiving_Summary | OK |
| E Invoice Validation | E_Invoice_Validation.E_Invoice_Validation | OK |
| Dabur Sale Fromat HA | Dabur_SaleFromat_HA.Dabur_SaleFromat_HA | OK |
| Password Breaker | Break_Password.PasswordBreaker | OK |
| Stkreport.Wisdom_Stkreport | Wisdom_Stkreport.Wisdom_Stkreport | OK |
| Convert Selected Cells | ConvertDatesAndNumbers.ConvertSelectedCells | OK |
| Combine Header And Clean | CombineHeaderAndClean.CombineHeaderAndClean | OK |
| Marico Hide Unwanted Columns | Marico_HideUnwantedColumns.Marico_HideUnwantedColumns | OK |
| Flipkart PO Formatter | Flipkart_PO_formatter.Flipkart_PO_formatter | OK |
| Blinkit PO Format | BlinkitPOFormatter.FormatBlinkitPO | OK |

## Group: DMS TATA
| Button label | Macro | Status |
|---|---|---|
| Customer Wise Sale Detail DMS | DMS_REP.DMS_REP | ? |
| DMS HSN Summary | DMS_HSN_SUMMARY.DMS_HSN_SUMMARY | OK |
| DMS Stock Statement | DMS_STK.DMS_STK | OK |
| E Invoice Format DMS MAX | E_invoice_Format_DMS_MAX.E_invoice_Format_DMS_MAX | OK |

## Group: Reliance
| Button label | Macro | Status |
|---|---|---|
| SRM Po Excel | RIL_SRM_Excel_Filter_Data.RIL_SRM_Excel_Filter_Data | ? |
| Print_Setup | Print_Setup_RIL.Print_Setup | OK |
| RIL Po To Excel | RIL_PO_TO_EXCEL.RIL_PO_TO_EXCEL | OK |
| Process For Amendment | RIL_Process_For_Amendment.RIL_Process_For_Amendment | OK (opens UserForm1) |
| Insert Brand Name ASN | RIL_Insert_Brand_Name_ASN.RIL_Insert_Brand_Name_ASN | OK |
| RIL Appointment | RIL_Appointment_Auto.RIL_Appointment | OK (single button, old duplicate removed) |

## Dropped (your decision)
Sale Comp, PDF to Excel, Random Distribute Items.

## Called internally, no button needed
RIL_PO_TO_EXCEL_PROCESS1, RIL_CLEAN_DATA1, RIL_Excel_Amendment_Upload, RIL__Excel_Amendment_Email (from UserForm1), Macro3, Mail_Selection_Range_Outlook_Body.
These are called with `Run "Name"`, so their names must not change.

## Macros in the file with no button (decide: add a button, or leave out)
| Macro | Module | What it seems to do |
|---|---|---|
| CleanPOData | Hands_on_Trade_PO_Clean | Cleans Hands-on-Trade PO data |
| VMM_Store_Rename | VMM_Store_Rename | Renames store names (old/new columns) |
| Macro1 | Wipro_Ecom_Report | Wipro Ecom report formatting |
| Macro2 | Module1 | Autofit and hide column B |
| SplitData | Module3 | Splits data into sheets by column value |
| ConvertTextToNumber | ConvertTextToNumber | Text to number |
| ConvertToNumber / ConvertToText | ConvertDatesAndNumbers | Variants of Convert Selected Cells |
| ShowAllColumns | Marico_HideUnwantedColumns | Undo for the Marico hide |
| SpellNumber | SpellNumber | Cell formula (UDF), no button needed |
