(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Aus (" + v + ")";
    case 1:  return "Heizen (" + v + ")";
    case 2:  return "Kühlen (" + v + ")";
    case 3:  return "Taupunkt (" + v + ")";
    case 4:  return "Pumpenwartungslauf (" + v + ")";
    case 5:  return "Frostschutz (" + v + ")";
    case 6:  return "Handbetrieb (" + v + ")";
    case 7:  return "Testmodus (" + v + ")";
    case 8:  return "Initialisierung (" + v + ")";
    case 9:  return "Sicherheitszustand (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)