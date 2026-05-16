(function(input) {
  var v = parseInt(input);
  if (v === 1) return "Ja";
  if (v === 0) return "Nein";
  return "Unbekannt (" + input + ")";
})(input)