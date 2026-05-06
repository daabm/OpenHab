(function(input) {
  var v = parseFloat(input);
  if (isNaN(v)) return input;
  return (v / 10.0).toFixed(1);
})(input)