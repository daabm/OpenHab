(function(input) {
  var v = parseFloat(input);
  if (isNaN(v)) return "OFF";
  return v > 0 ? "ON" : "OFF";
})(input)