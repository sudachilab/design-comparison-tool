const OVERVIEW_NUMBER_PATTERN = /\b[A-Z]\d+(?:[-_.][A-Z0-9]+)*\b/gi;

export function extractOverviewNumbers(value) {
  const matches = String(value).match(OVERVIEW_NUMBER_PATTERN) || [];
  return [...new Set(matches.map((match) => match.toUpperCase()))];
}

export function getOverviewColor(overviewNumber) {
  let hash = 0;

  for (const character of overviewNumber) {
    hash = (hash * 31 + character.charCodeAt(0)) % 360;
  }

  return `hsl(${hash} 78% 82%)`;
}
