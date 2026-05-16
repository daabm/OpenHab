(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Standby (" + v + ")";
    case 1:  return "Alarm (" + v + ")";
    case 2:  return "TWE (" + v + ")";
    case 3:  return "Kühlen (" + v + ")";
    case 4:  return "Heizen (" + v + ")";
    case 5:  return "Abtauung (" + v + ")";
    case 6:  return "Vorbereitung (" + v + ")";
    case 7:  return "Blockiert (" + v + ")";
    case 8:  return "EVU Sperre (" + v + ")";
    case 9:  return "Nicht verfügbar (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)