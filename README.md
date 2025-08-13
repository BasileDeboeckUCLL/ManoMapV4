# ManoMapV4 – EasyHRM

Automatische patroonanalyse in colonic high-resolution manometry (cHRM) data.  

🎯 **Doel**:  
Detecteren van anterograde, retrograde en simultane contractiele patronen, met export naar Excel voor medisch onderzoek en besluitvorming.

## 🔧 Componenten
- `export.py`:  
  Verwerkt patroondata en exporteert analyse met o.a. kleurcodering voor HAPCs (groen) en HARPCs (rood).
- `exportToExcelScreen.py`:  
  Gebruikersinterface met sliders voor instelbare drempels (mmHg, sensorlengte) en exports.

## ✅ Belangrijkste functionaliteiten
- Automatische herkenning van richtingen (a/r/s)
- Drempel- en sensorinstellingen via UI
- Statistische tabellen met kleurcodering
- Volledige `.xlsx` export

## 📂 Structuur
Zorg dat detectiebestanden worden verwerkt vóór export. Analysebestanden krijgen een `Statistics` tab met ingekleurde rijen.