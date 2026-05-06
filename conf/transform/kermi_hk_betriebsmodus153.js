(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Aus (" + v + ")";
    case 1:  return "Heizen (" + v + ")";
    case 2:  return "Kühlen (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)