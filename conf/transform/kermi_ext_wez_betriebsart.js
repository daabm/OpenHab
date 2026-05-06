(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Auto (" + v + ")";
    case 1:  return "Nur Wärmepumpe (" + v + ")";
    case 2:  return "Beide (" + v + ")";
    case 3:  return "Sekundärer Wärmeerzeuger (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)