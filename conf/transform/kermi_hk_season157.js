(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Auto (" + v + ")";
    case 1:  return "Heizen (" + v + ")";
    case 2:  return "Kühlen (" + v + ")";
    case 3:  return "Aus (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)