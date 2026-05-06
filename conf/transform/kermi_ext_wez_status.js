(function(input) {
  var v = parseInt(input);
  switch (v) {
    case 0:   return "Keine Anforderung (" + v + ")";
    case 100: return "Anforderung (" + v + ")";
    case 200: return "Bereitschaft Auto Parallel (" + v + ")";
    case 201: return "Bereitschaft Auto Alternativ (" + v + ")";
    case 204: return "Bereitschaft wegen Störung (" + v + ")";
    case 205: return "Bereitschaft Handbetrieb Parallel (" + v + ")";
    case 206: return "Bereitschaft wegen Handbetrieb Parallel (" + v + ")";
    case 207: return "Bereitschaft EVU Sperre (" + v + ")";
    case 300: return "Anforderung Auto Parallel (" + v + ")";
    case 301: return "Anforderung Auto Alternativ (" + v + ")";
    case 304: return "Anforderung wegen Störung (" + v + ")";
    case 305: return "Anforderung Handbetrieb Parallel (" + v + ")";
    case 306: return "Anforderung wegen Handbetrieb Parallel (" + v + ")";
    case 307: return "Anforderung EVU Sperre (" + v + ")";
    default:  return "Unbekannt (" + v + ")";
  }
})(input)