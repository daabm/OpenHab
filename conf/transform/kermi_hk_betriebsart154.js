(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Auto (" + v + ")";
    case 1:  return "Heizen (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)