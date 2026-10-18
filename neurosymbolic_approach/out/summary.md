# Pipeline run

| family | docs | abstained | values in key | correct | wrong | blank/missed | score |
|---|---|---|---|---|---|---|---|
| f00_demo | 4 | 1 | 1776 | 444 | 0 | 1332 | 444 |
| f01_net_spiegel | 8 | 2 | 2296 | 1768 | 0 | 528 | 1768 |
| f02_asset_spiegel | 7 | 2 | 8118 | 5796 | 0 | 2322 | 5796 |
| f03_block_spiegel | 7 | 7 | 5642 | 0 | 0 | 5642 | 0 |
| f04_verzeichnis | 8 | 2 | 1366 | 1042 | 0 | 324 | 1042 |
| f05_sap_class_gitter | 8 | 7 | 3204 | 432 | 0 | 2772 | 432 |
| f06_sap_asset_gitter | 7 | 3 | 5082 | 2888 | 0 | 2194 | 2888 |
| f07_sap_class_summary | 7 | 2 | 1071 | 794 | 0 | 277 | 794 |
| f08_excel_spiegel_2y | 8 | 2 | 1760 | 1320 | 0 | 440 | 1320 |
| f09_excel_register | 8 | 2 | 6072 | 4409 | 0 | 1663 | 4409 |
| **all** | 72 | 30 | 36387 | 18893 | 0 | 17494 | **18893** |

| doc | format | mode | rows | correct | wrong | blank/missed | score | note |
|---|---|---|---|---|---|---|---|---|
| f00_demo_01 | xlsx | line | 37 | 444 | 0 | 0 | 444 |  |
| f00_demo_02 | pdf | line | 37 | 0 | 0 | 444 | 0 |  |
| f00_demo_03 | image-pdf | none | 0 | 0 | 0 | 444 | 0 | no text layer: needs OCR (not wired in yet) |
| f00_demo_04 | pdf | line | 37 | 0 | 0 | 444 | 0 |  |
| f01_net_spiegel_01 | image-pdf | none | 0 | 0 | 0 | 264 | 0 | no text layer: needs OCR (not wired in yet) |
| f01_net_spiegel_02 | pdf | line | 33 | 264 | 0 | 0 | 264 |  |
| f01_net_spiegel_03 | pdf | line | 33 | 264 | 0 | 0 | 264 |  |
| f01_net_spiegel_04 | pdf | line | 23 | 184 | 0 | 0 | 184 |  |
| f01_net_spiegel_05 | image-pdf | none | 0 | 0 | 0 | 264 | 0 | no text layer: needs OCR (not wired in yet) |
| f01_net_spiegel_06 | pdf | line | 33 | 264 | 0 | 0 | 264 |  |
| f01_net_spiegel_07 | xlsx | line | 33 | 264 | 0 | 0 | 264 |  |
| f01_net_spiegel_08 | pdf | line | 66 | 528 | 0 | 0 | 528 |  |
| f02_asset_spiegel_01 | pdf | line | 112 | 1008 | 0 | 0 | 1008 |  |
| f02_asset_spiegel_02 | pdf | line | 137 | 1233 | 0 | 0 | 1233 |  |
| f02_asset_spiegel_03 | xlsx | line | 135 | 1215 | 0 | 0 | 1215 |  |
| f02_asset_spiegel_04 | image-pdf | none | 0 | 0 | 0 | 1107 | 0 | no text layer: needs OCR (not wired in yet) |
| f02_asset_spiegel_05 | pdf | line | 121 | 1089 | 0 | 0 | 1089 |  |
| f02_asset_spiegel_06 | image-pdf | none | 0 | 0 | 0 | 1215 | 0 | no text layer: needs OCR (not wired in yet) |
| f02_asset_spiegel_07 | pdf | line | 139 | 1251 | 0 | 0 | 1251 |  |
| f03_block_spiegel_01 | pdf | block | 0 | 0 | 0 | 793 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f03_block_spiegel_02 | xlsx | block | 0 | 0 | 0 | 819 | 0 | multi-line records (one asset/class = several rows): block reader not built yet |
| f03_block_spiegel_03 | image-pdf | none | 0 | 0 | 0 | 767 | 0 | no text layer: needs OCR (not wired in yet) |
| f03_block_spiegel_04 | image-pdf | none | 0 | 0 | 0 | 910 | 0 | no text layer: needs OCR (not wired in yet) |
| f03_block_spiegel_05 | pdf | block | 0 | 0 | 0 | 650 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f03_block_spiegel_06 | pdf | block | 0 | 0 | 0 | 936 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f03_block_spiegel_07 | pdf | block | 0 | 0 | 0 | 767 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f04_verzeichnis_01 | pdf | line | 84 | 168 | 0 | 0 | 168 |  |
| f04_verzeichnis_02 | pdf | line | 85 | 170 | 0 | 0 | 170 |  |
| f04_verzeichnis_03 | xlsx | line | 95 | 190 | 0 | 0 | 190 |  |
| f04_verzeichnis_04 | image-pdf | none | 0 | 0 | 0 | 154 | 0 | no text layer: needs OCR (not wired in yet) |
| f04_verzeichnis_05 | image-pdf | none | 0 | 0 | 0 | 170 | 0 | no text layer: needs OCR (not wired in yet) |
| f04_verzeichnis_06 | pdf | line | 82 | 164 | 0 | 0 | 164 |  |
| f04_verzeichnis_07 | pdf | line | 93 | 186 | 0 | 0 | 186 |  |
| f04_verzeichnis_08 | pdf | line | 82 | 164 | 0 | 0 | 164 |  |
| f05_sap_class_gitter_01 | pdf | block | 0 | 0 | 0 | 396 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f05_sap_class_gitter_02 | xlsx | block | 0 | 0 | 0 | 396 | 0 | multi-line records (one asset/class = several rows): block reader not built yet |
| f05_sap_class_gitter_03 | image-pdf | none | 0 | 0 | 0 | 360 | 0 | no text layer: needs OCR (not wired in yet) |
| f05_sap_class_gitter_04 | image-pdf | none | 0 | 0 | 0 | 432 | 0 | no text layer: needs OCR (not wired in yet) |
| f05_sap_class_gitter_05 | pdf | block | 0 | 0 | 0 | 408 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f05_sap_class_gitter_06 | pdf | block | 0 | 0 | 0 | 384 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f05_sap_class_gitter_07 | pdf | block | 0 | 0 | 0 | 396 | 0 | multi-line records (one asset/class = several printed lines): block reader not built yet |
| f05_sap_class_gitter_08 | pdf | line | 36 | 432 | 0 | 0 | 432 |  |
| f06_sap_asset_gitter_01 | image-pdf | none | 0 | 0 | 0 | 968 | 0 | no text layer: needs OCR (not wired in yet) |
| f06_sap_asset_gitter_02 | xlsx | line | 88 | 968 | 0 | 0 | 968 |  |
| f06_sap_asset_gitter_03 | image-pdf | none | 0 | 0 | 0 | 550 | 0 | no text layer: needs OCR (not wired in yet) |
| f06_sap_asset_gitter_04 | pdf | line | 67 | 737 | 0 | 0 | 737 |  |
| f06_sap_asset_gitter_05 | image-pdf | none | 0 | 0 | 0 | 627 | 0 | no text layer: needs OCR (not wired in yet) |
| f06_sap_asset_gitter_06 | pdf | line | 49 | 490 | 0 | 49 | 490 |  |
| f06_sap_asset_gitter_07 | pdf | line | 63 | 693 | 0 | 0 | 693 |  |
| f07_sap_class_summary_01 | pdf | line | 18 | 126 | 0 | 0 | 126 |  |
| f07_sap_class_summary_02 | xlsx | line | 18 | 126 | 0 | 0 | 126 |  |
| f07_sap_class_summary_03 | image-pdf | none | 0 | 0 | 0 | 140 | 0 | no text layer: needs OCR (not wired in yet) |
| f07_sap_class_summary_04 | image-pdf | none | 0 | 0 | 0 | 133 | 0 | no text layer: needs OCR (not wired in yet) |
| f07_sap_class_summary_05 | pdf | line | 40 | 280 | 0 | 0 | 280 |  |
| f07_sap_class_summary_06 | pdf | line | 18 | 122 | 0 | 4 | 122 |  |
| f07_sap_class_summary_07 | pdf | line | 20 | 140 | 0 | 0 | 140 |  |
| f08_excel_spiegel_2y_01 | xlsx | line | 20 | 220 | 0 | 0 | 220 |  |
| f08_excel_spiegel_2y_02 | pdf | line | 20 | 220 | 0 | 0 | 220 |  |
| f08_excel_spiegel_2y_03 | image-pdf | none | 0 | 0 | 0 | 220 | 0 | no text layer: needs OCR (not wired in yet) |
| f08_excel_spiegel_2y_04 | image-pdf | none | 0 | 0 | 0 | 220 | 0 | no text layer: needs OCR (not wired in yet) |
| f08_excel_spiegel_2y_05 | xlsx | line | 20 | 220 | 0 | 0 | 220 |  |
| f08_excel_spiegel_2y_06 | xlsx | line | 20 | 220 | 0 | 0 | 220 |  |
| f08_excel_spiegel_2y_07 | xlsx | line | 20 | 220 | 0 | 0 | 220 |  |
| f08_excel_spiegel_2y_08 | xlsx | line | 20 | 220 | 0 | 0 | 220 |  |
| f09_excel_register_01 | xlsx | line | 74 | 888 | 0 | 0 | 888 |  |
| f09_excel_register_02 | pdf | line | 62 | 744 | 0 | 0 | 744 |  |
| f09_excel_register_03 | image-pdf | none | 0 | 0 | 0 | 900 | 0 | no text layer: needs OCR (not wired in yet) |
| f09_excel_register_04 | image-pdf | none | 0 | 0 | 0 | 756 | 0 | no text layer: needs OCR (not wired in yet) |
| f09_excel_register_05 | xlsx | line | 66 | 785 | 0 | 7 | 785 |  |
| f09_excel_register_06 | xlsx | line | 46 | 552 | 0 | 0 | 552 |  |
| f09_excel_register_07 | xlsx | line | 59 | 708 | 0 | 0 | 708 |  |
| f09_excel_register_08 | xlsx | line | 61 | 732 | 0 | 0 | 732 |  |

## Detection of planted problems (documents with a text layer that were read)

| type | hit | missed | false alarm |
|---|---|---|---|
| carry_over | 1 | 0 | 0 |
| column_cut_off | 1 | 0 | 0 |
| glued_numbers | 5 | 0 | 0 |
| header_only_on_first_page | 1 | 0 | 0 |
| hidden_column | 1 | 0 | 0 |
| multi_year | 8 | 0 | 0 |
| noise_page_blank | 1 | 0 | 0 |
| noise_page_cover | 2 | 0 | 0 |
| noise_page_unrelated | 1 | 0 | 0 |
| non_tying_subtotal | 7 | 1 | 1 |
| numbers_as_text | 2 | 0 | 0 |
| teur_rounding | 1 | 0 | 0 |
| transfer_breaks_cost_law | 1 | 0 | 0 |
| unrelated_table | 1 | 0 | 0 |
| wrapped_description | 3 | 0 | 3 |
