 (function(input) {
  var v = parseInt(input);
  if (v === 1) return "Aktiv";
  if (v === 0) return "Nicht aktiv";
  return "Unbekannt (" + input + ")";
})(input)