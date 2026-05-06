(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:  return "Off (" + v + ")";
    case 1:  return "Eco (" + v + ")";
    case 2:  return "Normal (" + v + ")";
    case 3:  return "Comfort (" + v + ")";
    case 4:  return "Benutzerdefiniert (" + v + ")";
    default: return "Unbekannt (" + v + ")";
  }
})(input)