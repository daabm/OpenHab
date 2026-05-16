(function(input) {
  var v = parseFloat(input);
  if (isNaN(v)) return input;
  return Math.round(v * 10).toString();
})(input)