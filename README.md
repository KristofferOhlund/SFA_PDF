# SFA_PDF
Avläsning armeringsförteckning, 

Läser samtliga armeringsförteckningar utifrån en sökväg mha regex.
Kontrollera PDF, vissa PDF får en output sträng med "TOTALVIKT KG 0 VIKT"
Andra PDF har output "VIKTARMERINGSFÖRTECKNING"
Beroende på PDF bör alltså regex eventuellt justeras

Du kan summera totalvikten på två sätt.
Alternativ 1 - per sida
1. Justera parameter till Flatten=False för funktion read_from_path samt write_excel_summary
2. Kontrollera regex typ
3. Hämta vikt från fil samt resp. sida. 
4. Skapar excelrapport med 3 kolumner, FIL, SIDA; VIKT

Alternativ 2: Flatten 
1. Justera parameter till Flatten=True för funktion read_from_path samt write_excel_summary
2. Kontrollera regex typ
3. Hämtar vikt och slår ihop varje sida till en totalvikten
4. Skapar excelrapport med 2 kolumner, FIL, TOTALVIKT

